import std/[os, strformat, strutils]
import parsetoml

import ../../core/logging

const
  DEFAULT_INSPECTOR_STATE_FILENAME* = "uia_inspector_state.toml"

type
  InspectorState* = object
    leftWidth*: int
    middleWidth*: int
    infoHeight*: int
    propertiesHeight*: int
    highlightColor*: string
    highlightFollow*: bool
    filterVisible*: bool
    filterTitle*: bool
    filterActivate*: bool

proc defaultInspectorState*(): InspectorState =
  InspectorState(
    leftWidth: 320,
    middleWidth: 340,
    infoHeight: 180,
    propertiesHeight: 220,
    highlightColor: "#ff0000",
    highlightFollow: false,
    filterVisible: true,
    filterTitle: true,
    filterActivate: true
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

  let layout = root.getOrDefault("layout")
  if not layout.isNil and layout.kind == TomlValueKind.Table:
    result.leftWidth = getInt(layout.getOrDefault("left_width"), result.leftWidth)
    result.middleWidth = getInt(layout.getOrDefault("middle_width"), result.middleWidth)
    result.infoHeight = getInt(layout.getOrDefault("info_height"), result.infoHeight)
    result.propertiesHeight = getInt(layout.getOrDefault("properties_height"),
      result.propertiesHeight)
    # Backwards compatibility with earlier single-split layouts.
    if result.leftWidth <= 0:
      let legacyWidth = getInt(layout.getOrDefault("sash_width"), 0)
      if legacyWidth > 0:
        result.leftWidth = legacyWidth

  let display = root.getOrDefault("display")
  if not display.isNil and display.kind == TomlValueKind.Table:
    let parsed = getStr(display.getOrDefault("highlight_color"), result.highlightColor)
    if parsed.len > 0:
      result.highlightColor = parsed
    result.highlightFollow = getBool(display.getOrDefault("highlight_follow"),
      result.highlightFollow)

  let filters = root.getOrDefault("filters")
  if not filters.isNil and filters.kind == TomlValueKind.Table:
    result.filterVisible = getBool(filters.getOrDefault("visible"), result.filterVisible)
    result.filterTitle = getBool(filters.getOrDefault("title"), result.filterTitle)
    result.filterActivate = getBool(filters.getOrDefault("activate"), result.filterActivate)

proc saveInspectorState*(statePath: string, state: InspectorState, logger: Logger = nil) =
  ## Persist inspector layout and display preferences to disk.
  var lines: seq[string] = @[]
  lines.add("[layout]")
  lines.add(&"left_width = {state.leftWidth}")
  lines.add(&"middle_width = {state.middleWidth}")
  lines.add(&"info_height = {state.infoHeight}")
  lines.add(&"properties_height = {state.propertiesHeight}")
  lines.add("")
  lines.add("[display]")
  lines.add(&"highlight_color = \"{state.highlightColor}\"")
  lines.add(&"highlight_follow = {state.highlightFollow}")
  lines.add("")
  lines.add("[filters]")
  lines.add(&"visible = {state.filterVisible}")
  lines.add(&"title = {state.filterTitle}")
  lines.add(&"activate = {state.filterActivate}")
  lines.add("")

  try:
    writeFile(statePath, lines.join("\n"))
    if logger != nil:
      logger.info("Saved inspector state", [("path", statePath)])
  except CatchableError as exc:
    if logger != nil:
      logger.error("Failed to save inspector state",
        [("path", statePath), ("error", exc.msg)])
