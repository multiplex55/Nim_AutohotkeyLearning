# config_loader.nim
## Loads TOML config for the hotkey runner.
##
## Depends on:
##   - parsetoml  (nimble install parsetoml)
##   - logging.nim (your custom logging module)
##
## Expected TOML shape (example):
##
## [logging]
## level = "info"        # trace | debug | info | warn | error
## structured = true
##
## [[hotkeys]]
## keys   = "Ctrl+Alt+H"
## action = "paste_template"
##
##   [hotkeys.params]
##   template = "Hello, world"
##
##   [[hotkeys.sequence]]
##   delay_ms = 500
##   action   = "paste_template"
##
##     [hotkeys.sequence.params]
##     template = "Step 2"

import std/[options, tables, strutils]
import parsetoml
import ../core/logging

# ----- Public types --------------------------------------------------------

type
  StepConfig* = object
    delayMs*: int
    action*: string
    params*: Table[string, string]

  WindowTarget* = object
    name*: string
    title*: Option[string]
    titleContains*: Option[string]
    className*: Option[string]
    processName*: Option[string]
    storedHwnd*: Option[int]

  HotkeyConfig* = object
    enabled*: bool
    keys*: string
    action*: string
    params*: Table[string, string]
    target*: string
    uiaAction*: string
    uiaParams*: Table[string, string]
    delayMs*: Option[int]
    repeatMs*: Option[int]
    sequence*: seq[StepConfig]

  ConfigResult* = object
    loggingLevel*: Option[LogLevel]
    structuredLogs*: bool
    hotkeys*: seq[HotkeyConfig]
    windowTargets*: Table[string, WindowTarget]

# ----- Internal helpers ----------------------------------------------------

proc parseLogLevel(levelStr: string): Option[LogLevel] =
  let s = levelStr.toLowerAscii()
  case s
  of "trace":
    result = some(llTrace)
  of "debug":
    result = some(llDebug)
  of "info":
    result = some(llInfo)
  of "warn", "warning":
    result = some(llWarn)
  of "error":
    result = some(llError)
  else:
    result = none(LogLevel)

proc toParams(node: TomlValueRef): Table[string, string] =
  ## Convert a TOML table value into a simple string->string map.
  result = initTable[string, string]()
  if node.isNil or node.kind != TomlValueKind.Table:
    return

  let tbl = getTable(node)  # TomlTableRef
  if tbl.isNil:
    return

  for key, valRef in tbl[]:
    if valRef.isNil: continue
    case valRef.kind
    of TomlValueKind.String:
      result[key] = valRef.stringVal
    of TomlValueKind.Int:
      result[key] = $valRef.intVal
    of TomlValueKind.Float:
      result[key] = $valRef.floatVal
    of TomlValueKind.Bool:
      result[key] = $valRef.boolVal
    else:
      # Datetime/Date/Time/Array/Table etc. are ignored for params
      discard

proc parseWindowTarget(name: string, node: TomlValueRef, logger: Logger): WindowTarget =
  ## Extract a WindowTarget from a TOML table.
  result = WindowTarget(
    name: name,
    title: none(string),
    titleContains: none(string),
    className: none(string),
    processName: none(string),
    storedHwnd: none(int)
  )

  if node.isNil or node.kind != TomlValueKind.Table:
    logger.warn("Window target is not a table, skipping", [("name", name)])
    return

  let title = getStr(node{"title"}, "")
  if title.len > 0:
    result.title = some(title)

  let titleContains = getStr(node{"title_contains"}, "")
  if titleContains.len > 0:
    result.titleContains = some(titleContains)

  let className = getStr(node{"class"}, "")
  if className.len > 0:
    result.className = some(className)

  let processName = getStr(node{"process"}, "")
  if processName.len > 0:
    result.processName = some(processName)

  if parsetoml.hasKey(node, "hwnd"):
    result.storedHwnd = some(getInt(node{"hwnd"}))

# ----- Public API ----------------------------------------------------------

proc loadConfig*(path: string; logger: Logger): ConfigResult =
  ## Load configuration from a TOML file at `path`.
  ## Returns a ConfigResult with logging options and hotkey definitions.
  logger.info("Loading config", [("path", path)])

  let root = parseFile(path)  # TomlValueRef

  result.windowTargets = initTable[string, WindowTarget]()
  result.hotkeys = @[]

  if root.isNil or root.kind != TomlValueKind.Table:
    logger.error("Config root is not a TOML table", [])
    return

  # ----- Logging section ---------------------------------------------------
  let loggingNode = root{"logging"}
  if not loggingNode.isNil:
    let levelStr = getStr(loggingNode{"level"}, "")
    if levelStr.len > 0:
      let lvlOpt = parseLogLevel(levelStr)
      if lvlOpt.isSome:
        result.loggingLevel = lvlOpt
      else:
        logger.warn("Unknown logging level in config, using default", [("level", levelStr)])

    result.structuredLogs = getBool(loggingNode{"structured"}, false)
  else:
    result.structuredLogs = false

  # ----- Window targets ----------------------------------------------------
  let windowTargetsNode = root{"window_targets"}
  if not windowTargetsNode.isNil:
    case windowTargetsNode.kind
    of TomlValueKind.Array:
      for entry in getElems(windowTargetsNode):
        if entry.isNil or entry.kind != TomlValueKind.Table:
          continue
        let name = getStr(entry{"name"}, "")
        if name.len == 0:
          logger.warn("Window target missing name, skipping", [])
          continue
        let target = parseWindowTarget(name, entry, logger)
        result.windowTargets[name] = target
    of TomlValueKind.Table:
      let tbl = getTable(windowTargetsNode)
      for name, targetNode in tbl[]:
        if targetNode.isNil or targetNode.kind != TomlValueKind.Table:
          continue
        let target = parseWindowTarget(name, targetNode, logger)
        result.windowTargets[name] = target
    else:
      logger.warn("window_targets must be a table or array of tables", [])

  # ----- Hotkeys array -----------------------------------------------------
  let hotkeysNode = root{"hotkeys"}

  if hotkeysNode.isNil or hotkeysNode.kind != TomlValueKind.Array:
    logger.warn("No [[hotkeys]] array found in config", [])
  else:
    for entry in getElems(hotkeysNode):
      if entry.isNil or entry.kind != TomlValueKind.Table:
        continue

      var hk: HotkeyConfig
      hk.enabled = getBool(entry{"enabled"}, true)
      hk.keys   = getStr(entry{"keys"}, "")
      hk.action = getStr(entry{"action"}, "")
      hk.params = initTable[string, string]()
      hk.target = getStr(entry{"target"}, "")
      hk.uiaAction = getStr(entry{"uia_action"}, "")
      hk.uiaParams = initTable[string, string]()

      if hk.keys.len == 0 or (hk.action.len == 0 and hk.uiaAction.len == 0):
        logger.warn("Hotkey entry missing 'keys' or action/uia_action, skipping", [])
        continue

      # top-level params
      let paramsNode = entry{"params"}
      if not paramsNode.isNil and paramsNode.kind == TomlValueKind.Table:
        hk.params = toParams(paramsNode)

      # uia params
      let uiaParamsNode = entry{"uia_params"}
      if not uiaParamsNode.isNil and uiaParamsNode.kind == TomlValueKind.Table:
        hk.uiaParams = toParams(uiaParamsNode)

      # delay_ms / repeat_ms (optional)
      if parsetoml.hasKey(entry, "delay_ms"):
        hk.delayMs = some(getInt(entry{"delay_ms"}))
      else:
        hk.delayMs = none(int)

      if parsetoml.hasKey(entry, "repeat_ms"):
        hk.repeatMs = some(getInt(entry{"repeat_ms"}))
      else:
        hk.repeatMs = none(int)

      # sequence steps
      hk.sequence = @[]
      let seqNode = entry{"sequence"}
      if not seqNode.isNil and seqNode.kind == TomlValueKind.Array:
        for stepVal in getElems(seqNode):
          if stepVal.isNil or stepVal.kind != TomlValueKind.Table:
            continue

          var step: StepConfig
          step.delayMs = getInt(stepVal{"delay_ms"})
          step.action  = getStr(stepVal{"action"}, "")
          let stepParamsNode = stepVal{"params"}
          if not stepParamsNode.isNil and stepParamsNode.kind == TomlValueKind.Table:
            step.params = toParams(stepParamsNode)
          else:
            step.params = initTable[string, string]()

          hk.sequence.add(step)

      result.hotkeys.add(hk)

  logger.info("Config loaded", [("hotkey_count", $result.hotkeys.len)])
