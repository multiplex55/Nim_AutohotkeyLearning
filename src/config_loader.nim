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
import logging

# ----- Public types --------------------------------------------------------

type
  StepConfig* = object
    delayMs*: int
    action*: string
    params*: Table[string, string]

  HotkeyConfig* = object
    keys*: string
    action*: string
    params*: Table[string, string]
    delayMs*: Option[int]
    repeatMs*: Option[int]
    sequence*: seq[StepConfig]

  ConfigResult* = object
    loggingLevel*: Option[LogLevel]
    structuredLogs*: bool
    hotkeys*: seq[HotkeyConfig]

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

# ----- Public API ----------------------------------------------------------

proc loadConfig*(path: string; logger: Logger): ConfigResult =
  ## Load configuration from a TOML file at `path`.
  ## Returns a ConfigResult with logging options and hotkey definitions.
  logger.info("Loading config", [("path", path)])

  let root = parseFile(path)  # TomlValueRef

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

  # ----- Hotkeys array -----------------------------------------------------
  let hotkeysNode = root{"hotkeys"}

  if hotkeysNode.isNil or hotkeysNode.kind != TomlValueKind.Array:
    logger.warn("No [[hotkeys]] array found in config", [])
  else:
    for entry in getElems(hotkeysNode):
      if entry.isNil or entry.kind != TomlValueKind.Table:
        continue

      var hk: HotkeyConfig
      hk.keys   = getStr(entry{"keys"}, "")
      hk.action = getStr(entry{"action"}, "")
      hk.params = initTable[string, string]()

      if hk.keys.len == 0 or hk.action.len == 0:
        logger.warn("Hotkey entry missing 'keys' or 'action', skipping", [])
        continue

      # top-level params
      let paramsNode = entry{"params"}
      if not paramsNode.isNil and paramsNode.kind == TomlValueKind.Table:
        hk.params = toParams(paramsNode)

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
