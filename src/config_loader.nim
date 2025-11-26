import std/[options, tables]

when compiles do:
  import std/toml as toml
elif compiles do:
  import pkg/toml as toml
else:
  {.error: "No TOML parser available; install std/toml (Nim 2+) or pkg/toml".}

when declared(toml.TomlValueRef):
  type TomlValue = toml.TomlValueRef
elif declared(toml.TomlValue):
  type TomlValue = toml.TomlValue
else:
  {.error: "Unsupported TOML value type exposed by selected parser".}

import ./logging

type
  StepConfig* = object
    delayMs*: int
    action*: string
    params*: Table[string, string]

  HotkeyConfig* = object
    name*: string
    keys*: string
    action*: string
    params*: Table[string, string]
    delayMs*: Option[int]
    repeatMs*: Option[int]
    sequence*: seq[StepConfig]

  ConfigResult* = object
    loggingLevel*: Option[string]
    structuredLogs*: bool
    hotkeys*: seq[HotkeyConfig]

proc toParams(tbl: auto): Table[string, string] =
  result = initTable[string, string]()
  for key, val in tbl.pairs:
    case val.kind
    of toml.TomlKind.Int:
      result[key] = $val.intVal
    of toml.TomlKind.Float:
      result[key] = $val.floatVal
    of toml.TomlKind.Bool:
      result[key] = if val.boolVal: "true" else: "false"
    of toml.TomlKind.String:
      result[key] = val.stringVal
    else:
      discard

proc loadConfig*(path: string, logger: Logger = nil): ConfigResult =
  if logger != nil:
    logger.info("Loading config", [("path", path)])

  var parsed: TomlValue
  try:
    parsed = toml.parseFile(path)
  except CatchableError as e:
    if logger != nil:
      logger.error("Failed to parse config", [("error", e.msg)])
    return

  if "logging" in parsed and parsed["logging"].kind == toml.TomlKind.Table:
    let loggingTable = parsed["logging"].tableVal
    if "level" in loggingTable:
      result.loggingLevel = some(loggingTable["level"].stringVal)
    if "structured" in loggingTable:
      result.structuredLogs = loggingTable["structured"].boolVal

  if "hotkey" in parsed:
    let hkVal = parsed["hotkey"]
    if hkVal.kind == toml.TomlKind.Array and hkVal.arrayVal.len > 0:
      for entry in hkVal.arrayVal:
        if entry.kind != toml.TomlKind.Table:
          continue
        let tbl = entry.tableVal
        var cfg = HotkeyConfig(
          name: if "name" in tbl: tbl["name"].stringVal else: "",
          keys: if "keys" in tbl: tbl["keys"].stringVal else: "",
          action: if "action" in tbl: tbl["action"].stringVal else: "",
          params: initTable[string, string](),
          sequence: @[]
        )

        if "params" in tbl and tbl["params"].kind == toml.TomlKind.Table:
          cfg.params = toParams(tbl["params"].tableVal)

        if "delay_ms" in tbl and tbl["delay_ms"].kind == toml.TomlKind.Int:
          cfg.delayMs = some(int(tbl["delay_ms"].intVal))

        if "repeat_ms" in tbl and tbl["repeat_ms"].kind == toml.TomlKind.Int:
          cfg.repeatMs = some(int(tbl["repeat_ms"].intVal))

        if "sequence" in tbl and tbl["sequence"].kind == toml.TomlKind.Array:
          for stepVal in tbl["sequence"].arrayVal:
            if stepVal.kind != toml.TomlKind.Table:
              continue
            let stepTable = stepVal.tableVal
            var stepCfg = StepConfig(
              delayMs: if "delay_ms" in stepTable: stepTable["delay_ms"].intVal.int else: 0,
              action: if "action" in stepTable: stepTable["action"].stringVal else: "",
              params: initTable[string, string]()
            )
            if "params" in stepTable and stepTable["params"].kind == toml.TomlKind.Table:
              stepCfg.params = toParams(stepTable["params"].tableVal)
            cfg.sequence.add(stepCfg)

        result.hotkeys.add(cfg)

  if logger != nil:
    logger.info("Loaded hotkey entries", [("count", $result.hotkeys.len)])
