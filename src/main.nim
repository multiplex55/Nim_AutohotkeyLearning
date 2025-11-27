import std/[options, os, strformat, tables, times]

import ./core/[logging, runtime_context, scheduler, platform_backend]
import ./features/[actions, config_loader, key_parser, plugins, window_target_state]

when defined(windows):
  import ./platform/windows/backend as winBackend
  import ./features/win_automation/windows_helpers
  import ./features/uia/uia_plugin
else:
  import ./platform/linux/backend as linuxBackend

const DEFAULT_CONFIG = "examples/hotkeys.toml"

proc buildCallback(cfg: HotkeyConfig, registry: ActionRegistry, ctx: var RuntimeContext): HotkeyCallback =
  let actionName =
    if cfg.action.len > 0:
      cfg.action
    else:
      cfg.uiaAction

  var actionParams: Table[string, string]
  if cfg.action.len > 0:
    actionParams = cfg.params
  else:
    actionParams = cfg.uiaParams

  if cfg.target.len > 0 and not actionParams.hasKey("target"):
    actionParams["target"] = cfg.target

  let baseAction = registry.createAction(actionName, actionParams, ctx)

  # Sequence of actions with per-step delays
  if cfg.sequence.len > 0:
    var steps: seq[ScheduledStep] = @[]
    for step in cfg.sequence:
      let stepAction = registry.createAction(step.action, step.params, ctx)
      steps.add(
        ScheduledStep(
          delay: initDuration(milliseconds = step.delayMs),
          action: stepAction
        )
      )

    return proc() =
      if ctx.logger != nil:
        ctx.logger.info("Running sequence", [("hotkey", cfg.keys)])
      discard ctx.scheduler.scheduleSequence(steps)

  # Repeating task
  if cfg.repeatMs.isSome:
    return proc() =
      if ctx.logger != nil:
        ctx.logger.info(
          "Scheduling repeating task",
          [("hotkey", cfg.keys), ("interval", $cfg.repeatMs.get())]
        )
      discard ctx.scheduler.scheduleRepeat(
        initDuration(milliseconds = cfg.repeatMs.get()),
        baseAction
      )

  # One-shot delayed task
  if cfg.delayMs.isSome:
    return proc() =
      if ctx.logger != nil:
        ctx.logger.info(
          "Scheduling delayed task",
          [("hotkey", cfg.keys), ("delay", $cfg.delayMs.get())]
        )
      discard ctx.scheduler.scheduleOnce(
        initDuration(milliseconds = cfg.delayMs.get()),
        baseAction
      )

  # Immediate task
  return proc() =
    if ctx.logger != nil:
      ctx.logger.info("Executing immediate action", [("hotkey", cfg.keys)])
    baseAction()

proc registerConfiguredHotkeys*(
    config: ConfigResult,
    backend: PlatformBackend,
    registry: ActionRegistry,
    runtime: var RuntimeContext,
    clearExisting: bool = true
  ): int =
  let logger = runtime.logger

  if clearExisting:
    try:
      backend.clearHotkeys()
      if logger != nil:
        logger.debug("Cleared existing hotkeys before registration")
    except CatchableError as e:
      if logger != nil:
        logger.warn("Failed to clear existing hotkeys", [("error", e.msg)])

  for hk in config.hotkeys:
    if not hk.enabled:
      if logger != nil:
        logger.info("Hotkey disabled; skipping registration", [("keys", hk.keys)])
      continue

    let parsed = parseHotkeyString(hk.keys)
    if parsed.key == 0:
      logger.warn("Skipping hotkey with no key", [("keys", hk.keys)])
      continue

    try:
      let cb = buildCallback(hk, registry, runtime)
      discard backend.registerHotkey(parsed.modifiers, parsed.key, cb)
      logger.info(
        "Registered hotkey",
        [
          ("keys", hk.keys),
          ("action", if hk.action.len > 0: hk.action else: hk.uiaAction)
        ]
      )
      inc result
    except IOError as e:
      logger.error(
        "Failed to register hotkey",
        [("keys", hk.keys), ("error", e.msg)]
      )

proc setupHotkeys(configPath: string): bool =
  var logger = newLogger()
  var scheduler = newScheduler(logger)
  let backend: PlatformBackend =
    when defined(windows):
      winBackend.newWindowsBackend()
    else:
      linuxBackend.newLinuxBackend()
  let statePath = deriveStatePath(configPath)

  var registry = newActionRegistry(logger)
  registerBuiltinActions(registry)

  # Explicit type keeps nimsuggest from getting confused about fields like
  # loggingLevel / structuredLogs / hotkeys.
  let config: ConfigResult = loadConfig(configPath, logger)

  var targets = config.windowTargets
  loadWindowTargetState(statePath, targets, logger)

  var runtime = RuntimeContext(
    logger: logger,
    scheduler: scheduler,
    backend: backend,
    windowTargets: targets,
    windowTargetStatePath: some(statePath)
  )

  var pluginManager = newPluginManager(logger)
  when defined(windows):
    pluginManager.registerPlugin(newUiaPlugin(), registry, runtime)
    pluginManager.registerPlugin(newWindowsHelpers(), registry, runtime)

  if config.loggingLevel.isSome:
    logger.setLogLevel(config.loggingLevel.get())
  logger.structured = config.structuredLogs

  discard registerConfiguredHotkeys(config, backend, registry, runtime)

  echo "Entering message loop. Press configured exit hotkey to quit."
  backend.runMessageLoop(scheduler)
  pluginManager.shutdownPlugins(runtime)
  result = true

when isMainModule:
  let configPath =
    if paramCount() >= 1:
      paramStr(1)
    else:
      DEFAULT_CONFIG

  if not fileExists(configPath):
    echo &"Config file {configPath} not found."
  else:
    discard setupHotkeys(configPath)
