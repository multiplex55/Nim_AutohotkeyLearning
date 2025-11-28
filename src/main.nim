import std/[options, os, strformat, strutils, tables, times]

import ./core/[logging, runtime_context, scheduler, platform_backend]
import ./features/[actions, config_loader, key_parser, plugins, window_target_state]

import ./platform/windows/backend as winBackend
import ./features/win_automation/windows_helpers
import ./features/uia/uia_plugin

when defined(windows):
  import winim/lean
  import winim/inc/uiautomation

  import ./features/uia/uia
  import ./platform/windows/processes as winProcesses
  import ./platform/windows/windows as winWindows

const DEFAULT_CONFIG = "examples/hotkeys.toml"


when defined(windows):
  proc ensureHrOk(hr: HRESULT, ctx: string) =
    if FAILED(hr):
      raise newException(UiaError, fmt"{ctx} failed (0x{hr:X})")

  proc controlTypeName(typeId: int): string =
    case typeId
    of UIA_ButtonControlTypeId: "Button"
    of UIA_CalendarControlTypeId: "Calendar"
    of UIA_CheckBoxControlTypeId: "CheckBox"
    of UIA_ComboBoxControlTypeId: "ComboBox"
    of UIA_DataGridControlTypeId: "DataGrid"
    of UIA_DocumentControlTypeId: "Document"
    of UIA_EditControlTypeId: "Edit"
    of UIA_GroupControlTypeId: "Group"
    of UIA_HyperlinkControlTypeId: "Hyperlink"
    of UIA_ImageControlTypeId: "Image"
    of UIA_ListControlTypeId: "List"
    of UIA_ListItemControlTypeId: "ListItem"
    of UIA_MenuControlTypeId: "Menu"
    of UIA_MenuBarControlTypeId: "MenuBar"
    of UIA_MenuItemControlTypeId: "MenuItem"
    of UIA_PaneControlTypeId: "Pane"
    of UIA_ProgressBarControlTypeId: "ProgressBar"
    of UIA_RadioButtonControlTypeId: "RadioButton"
    of UIA_ScrollBarControlTypeId: "ScrollBar"
    of UIA_SplitButtonControlTypeId: "SplitButton"
    of UIA_StatusBarControlTypeId: "StatusBar"
    of UIA_TabControlTypeId: "Tab"
    of UIA_TabItemControlTypeId: "TabItem"
    of UIA_TextControlTypeId: "Text"
    of UIA_TitleBarControlTypeId: "TitleBar"
    of UIA_ToolBarControlTypeId: "ToolBar"
    of UIA_ToolTipControlTypeId: "ToolTip"
    of UIA_TreeControlTypeId: "Tree"
    of UIA_TreeItemControlTypeId: "TreeItem"
    of UIA_WindowControlTypeId: "Window"
    else: fmt"ControlType({typeId})"

  proc findNotepadWindow(): HWND =
    result = FindWindowW("Notepad", nil)
    if result != 0:
      return

    let untitled = winWindows.findWindowByTitleExact("Untitled - Notepad")
    if untitled != 0:
      result = HWND(untitled)

  proc logElementTree(uia: Uia, element: ptr IUIAutomationElement,
      walker: ptr IUIAutomationTreeWalker, depth, maxDepth: int,
      logger: Logger) =
    if element.isNil or depth > maxDepth:
      return

    let name = element.currentName()
    let ctrlType = controlTypeName(element.currentControlType())
    let hwnd = element.nativeWindowHandle()
    var fields: seq[(string, string)] = @[
      ("controlType", ctrlType),
      ("depth", $depth)
    ]
    if name.len > 0:
      fields.add(("name", name))
    if hwnd != 0:
      fields.add(("hwnd", fmt"0x{cast[uint](hwnd):X}"))

    let indent = "  ".repeat(depth)
    logger.info(indent & "- UIA element", fields)

    if depth == maxDepth:
      return

    var child: ptr IUIAutomationElement
    let hrFirst = walker.GetFirstChildElement(element, addr child)
    if FAILED(hrFirst):
      ensureHrOk(hrFirst, "GetFirstChildElement")
    if hrFirst == S_FALSE or child.isNil:
      return

    var current = child
    while current != nil:
      logElementTree(uia, current, walker, depth + 1, maxDepth, logger)

      var next: ptr IUIAutomationElement
      let hrNext = walker.GetNextSiblingElement(current, addr next)
      discard current.Release()
      if FAILED(hrNext):
        ensureHrOk(hrNext, "GetNextSiblingElement")
      if hrNext == S_FALSE:
        break
      current = next

  proc runUiaDemo(maxDepth: int = 3): int =
    var logger = newLogger()
    let existing = winProcesses.findProcessesByName("notepad.exe")
    if existing.len == 0:
      logger.info("Notepad not running; starting it")
      if not winProcesses.startProcessDetached("notepad.exe"):
        logger.error("Failed to start notepad.exe")
        return 1
      sleep(500.milliseconds)
    else:
      logger.info("Using existing Notepad instance", [("count", $existing.len)])

    var hwnd = findNotepadWindow()
    var attempts = 0
    while hwnd == 0 and attempts < 10:
      sleep(200.milliseconds)
      hwnd = findNotepadWindow()
      inc attempts

    if hwnd == 0:
      logger.error("Could not find a Notepad window. Is Notepad visible?")
      return 1

    logger.info(
      "Found Notepad window",
      [
        ("title", winWindows.getWindowTitle(hwnd)),
        ("hwnd", fmt"0x{cast[uint](hwnd):X}")
      ]
    )

    let uia = initUia()
    defer: uia.shutdown()

    let element = uia.fromWindowHandle(hwnd)
    if element.isNil:
      logger.error("Failed to obtain UIA element for Notepad window")
      return 1
    defer: discard element.Release()

    var walker: ptr IUIAutomationTreeWalker
    ensureHrOk(uia.automation.get_RawViewWalker(addr walker), "RawViewWalker")
    defer:
      if walker != nil:
        discard walker.Release()

    logger.info("Dumping Notepad UIA subtree", [("maxDepth", $maxDepth)])
    logElementTree(uia, element, walker, 0, maxDepth, logger)

    result = 0
else:
  proc runUiaDemo(maxDepth: int = 3): int =
    discard maxDepth
    echo "The --uia-demo flag is only available on Windows targets."
    1


proc buildCallback(cfg: HotkeyConfig, registry: ActionRegistry,
    ctx: var RuntimeContext): HotkeyCallback =
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

  # pull what we need from ctx *before* creating closures
  let logger = ctx.logger
  let sched = ctx.scheduler

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
      if logger != nil:
        logger.info("Running sequence", [("hotkey", cfg.keys)])
      discard sched.scheduleSequence(steps)

  # Repeating task
  if cfg.repeatMs.isSome:
    return proc() =
      if logger != nil:
        logger.info(
          "Scheduling repeating task",
          [("hotkey", cfg.keys), ("interval", $cfg.repeatMs.get())]
        )
      discard sched.scheduleRepeat(
        initDuration(milliseconds = cfg.repeatMs.get()),
        baseAction
      )

  # One-shot delayed task
  if cfg.delayMs.isSome:
    return proc() =
      if logger != nil:
        logger.info(
          "Scheduling delayed task",
          [("hotkey", cfg.keys), ("delay", $cfg.delayMs.get())]
        )
      discard sched.scheduleOnce(
        initDuration(milliseconds = cfg.delayMs.get()),
        baseAction
      )

  # Immediate task
  return proc() =
    if logger != nil:
      logger.info("Executing immediate action", [("hotkey", cfg.keys)])
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
  if paramCount() >= 1 and paramStr(1) == "--uia-demo":
    quit(runUiaDemo(4))

  let configPath =
    if paramCount() >= 1:
      paramStr(1)
    else:
      DEFAULT_CONFIG

  if not fileExists(configPath):
    echo &"Config file {configPath} not found."
  else:
    discard setupHotkeys(configPath)
