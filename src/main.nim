import std/[options, os, strformat, times]

import ./core/[logging, runtime_context, scheduler, platform_backend]
import ./features/[actions, config_loader, key_parser, plugins]

when defined(windows):
  import ./platform/windows/backend as winBackend
  import ./features/win_automation/windows_helpers
else:
  import ./platform/linux/backend as linuxBackend

const DEFAULT_CONFIG = "examples/hotkeys.toml"

proc buildCallback(cfg: HotkeyConfig, registry: ActionRegistry, ctx: RuntimeContext): HotkeyCallback =
  let baseAction = registry.createAction(cfg.action, cfg.params, ctx)

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

proc setupHotkeys(configPath: string): bool =
  var logger = newLogger()
  var scheduler = newScheduler(logger)
  let backend: PlatformBackend =
    when defined(windows):
      winBackend.newWindowsBackend()
    else:
      linuxBackend.newLinuxBackend()
  var registry = newActionRegistry(logger)
  registerBuiltinActions(registry)

  let runtime = RuntimeContext(logger: logger, scheduler: scheduler, backend: backend)

  var pluginManager = newPluginManager(logger)
  when defined(windows):
    pluginManager.registerPlugin(newWindowsHelpers(), registry, runtime)

  # Explicit type keeps nimsuggest from getting confused about fields like
  # loggingLevel / structuredLogs / hotkeys.
  let config: ConfigResult = loadConfig(configPath, logger)

  if config.loggingLevel.isSome:
    logger.setLogLevel(config.loggingLevel.get())
  logger.structured = config.structuredLogs

  for hk in config.hotkeys:
    let parsed = parseHotkeyString(hk.keys)
    if parsed.key == 0:
      logger.warn("Skipping hotkey with no key", [("keys", hk.keys)])
      continue

    try:
      let cb = buildCallback(hk, registry, runtime)
      discard backend.registerHotkey(parsed.modifiers, parsed.key, cb)
      logger.info(
        "Registered hotkey",
        [("keys", hk.keys), ("action", hk.action)]
      )
    except IOError as e:
      logger.error(
        "Failed to register hotkey",
        [("keys", hk.keys), ("error", e.msg)]
      )

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
