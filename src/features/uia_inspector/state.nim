import std/[options, os, strformat, strutils]
import parsetoml

import ../../core/logging

const
  DEFAULT_INSPECTOR_STATE_FILENAME* = "uia_inspector_state.toml"

type
  InspectorState* = object
    sashWidth*: int
    highlightColor*: string

proc defaultInspectorState*(): InspectorState =
  InspectorState(
    sashWidth: 420,
    highlightColor: "#ff0000"
  )

proc deriveInspectorStatePath*(configPath: string): string =
  ## Place the inspector state file next to the provided config/state file.
  let dir = splitPath(configPath).head
  joinPath(dir, DEFAULT_INSPECTOR_STATE_FILENAME)

proc loadInspectorState*(statePath: string, logger: Logger = nil): InspectorState =
  ## Load persisted inspector settings (splitter width, highlight color).
  result = defaultInspectorState()
  if not fileExists(statePath):
    return

  var root: TomlValueRef
  try:
    root = parseFile(statePath)
  except CatchableError as exc:
    if logger != nil:
      logger.warn("Failed to parse inspector state file",
        [("path", statePath), ("error", exc.msg)])
    return

  if root.isNil or root.kind != TomlValueKind.Table:
    if logger != nil:
      logger.warn("Inspector state file invalid; expected table root", [("path", statePath)])
    return

  let layout = root{"layout"}
  if not layout.isNil and layout.kind == TomlValueKind.Table:
    result.sashWidth = getInt(layout{"sash_width"}, result.sashWidth)

  let display = root{"display"}
  if not display.isNil and display.kind == TomlValueKind.Table:
    let parsed = getStr(display{"highlight_color"}, result.highlightColor)
    if parsed.len > 0:
      result.highlightColor = parsed

proc saveInspectorState*(statePath: string, state: InspectorState, logger: Logger = nil) =
  ## Persist inspector layout and display preferences to disk.
  var lines: seq[string] = @[]
  lines.add("[layout]")
  lines.add(&"sash_width = {state.sashWidth}")
  lines.add("")
  lines.add("[display]")
  lines.add(&"highlight_color = \"{state.highlightColor}\"")
  lines.add("")

  try:
    writeFile(statePath, lines.join("\n"))
    if logger != nil:
      logger.info("Saved inspector state", [("path", statePath)])
  except CatchableError as exc:
    if logger != nil:
      logger.error("Failed to save inspector state",
        [("path", statePath), ("error", exc.msg)])
