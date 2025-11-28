import std/[options, os, strutils, tables, strformat]
import parsetoml

import ../core/logging
import ../core/window_targets

const DEFAULT_STATE_FILENAME* = "window_targets_state.toml"

proc deriveStatePath*(configPath: string): string =
  ## Place the state file next to the primary config.
  let dir = splitPath(configPath).head
  result = joinPath(dir, DEFAULT_STATE_FILENAME)

proc validateHandle*(hwnd: int, logger: Logger): bool =
  if hwnd <= 0:
    if logger != nil:
      logger.warn("Invalid HWND found in state; skipping", [("hwnd", $hwnd)])
    return false
  true

proc loadWindowTargetState*(statePath: string, targets: var Table[string, WindowTarget], logger: Logger) =
  ## Merge any persisted HWNDs into the in-memory target map.
  if not fileExists(statePath):
    return

  var root: TomlValueRef
  try:
    root = parseFile(statePath)
  except CatchableError as e:
    if logger != nil:
      logger.warn("Failed to parse window target state file", [("path", statePath), ("error", e.msg)])
    return
  if root.isNil or root.kind != TomlValueKind.Table:
    if logger != nil:
      logger.warn("Window target state file invalid; expected table root", [("path", statePath)])
    return

  let node = root{"window_targets"}
  if node.isNil or node.kind != TomlValueKind.Table:
    if logger != nil:
      logger.warn("window_targets section missing in state file", [("path", statePath)])
    return

  for name, targetNode in getTable(node)[]:
    if targetNode.isNil or targetNode.kind != TomlValueKind.Table:
      continue
    if not parsetoml.hasKey(targetNode, "hwnd"):
      continue

    let hwnd = getInt(targetNode{"hwnd"})
    if not validateHandle(hwnd, logger):
      continue

    updateStoredHwnd(targets, name, hwnd, logger)

proc saveWindowTargetState*(statePath: string, targets: Table[string, WindowTarget], logger: Logger) =
  ## Persist HWNDs to disk so they survive restarts.
  var lines: seq[string] = @[]
  lines.add("[window_targets]")

  for name, target in targets.pairs():
    if target.storedHwnd.isNone:
      continue
    lines.add(&"[window_targets.{name}]")
    lines.add(&"hwnd = {target.storedHwnd.get()}")

  try:
    writeFile(statePath, lines.join("\n") & "\n")
    if logger != nil:
      logger.info("Saved window target state", [("path", statePath)])
  except CatchableError as e:
    if logger != nil:
      logger.error("Failed to save window target state", [("path", statePath), ("error", e.msg)])
