import std/[options, strutils, tables]

import ../core/logging
import ../core/runtime_context
import ../core/platform_backend
import ../core/scheduler
import ../core/window_targets
import ./window_target_state

type
  ActionFactory* = proc(params: Table[string, string],
      ctx: var RuntimeContext): TaskAction

  ActionRegistry* = ref object
    factories: Table[string, ActionFactory]
    logger: Logger

proc newActionRegistry*(logger: Logger = nil): ActionRegistry =
  ActionRegistry(factories: initTable[string, ActionFactory](), logger: logger)

proc registerAction*(registry: ActionRegistry, name: string,
    factory: ActionFactory) =
  registry.factories[name.toLowerAscii()] = factory
  if registry.logger != nil:
    registry.logger.debug("Registered action", [("name", name)])

proc createAction*(registry: ActionRegistry, name: string, params: Table[string,
    string], ctx: var RuntimeContext): TaskAction =
  let key = name.toLowerAscii()
  if key notin registry.factories:
    if registry.logger != nil:
      registry.logger.warn("Unknown action requested", [("action", name)])
    return proc() = discard
  registry.factories[key](params, ctx)

proc parseIntOpt(params: Table[string, string], key: string,
    default: int = 0): int =
  if key in params:
    try:
      result = parseInt(params[key])
    except ValueError:
      result = default
  else:
    result = default

proc parseBoolOpt(params: Table[string, string], key: string,
    default: bool = true): bool =
  if key in params:
    let value = params[key].toLowerAscii()
    return value in ["1", "true", "yes", "on"]
  default

proc registerBuiltinActions*(registry: ActionRegistry) =
  ## Built-in actions that don't require plugins.

  registry.registerAction("start_process", proc(params: Table[string, string],
      ctx: var RuntimeContext): TaskAction =
    let cmd = params.getOrDefault("command", "")
    let backend = ctx.backend
    let logger = ctx.logger
    return proc() =
      if backend.startProcessDetached(cmd):
        if logger != nil:
          logger.info("Started process", [("command", cmd)])
      else:
        if logger != nil:
          logger.error("Failed to start process", [("command", cmd)])
  )

  registry.registerAction("kill_process", proc(params: Table[string, string],
      ctx: var RuntimeContext): TaskAction =
    let name = params.getOrDefault("name", "")
    let backend = ctx.backend
    let logger = ctx.logger
    return proc() =
      let killed = backend.killProcessesByName(name)
      if logger != nil:
        logger.info("Kill process result", [("name", name), ("killed", $killed)])
  )

  registry.registerAction("send_text", proc(params: Table[string, string],
      ctx: var RuntimeContext): TaskAction =
    let msg = params.getOrDefault("text", "")
    let backend = ctx.backend
    let logger = ctx.logger
    return proc() =
      backend.sendText(msg)
      if logger != nil:
        logger.info("Sent text", [("text", msg)])
  )

  registry.registerAction("move_mouse", proc(params: Table[string, string],
      ctx: var RuntimeContext): TaskAction =
    let x = parseIntOpt(params, "x")
    let y = parseIntOpt(params, "y")
    let backend = ctx.backend
    let logger = ctx.logger
    return proc() =
      if backend.setMousePos(x, y):
        if logger != nil:
          logger.info("Mouse moved", [("x", $x), ("y", $y)])
      else:
        if logger != nil:
          logger.error("Failed to move mouse", [("x", $x), ("y", $y)])
  )



  registry.registerAction("left_click", proc(params: Table[string, string],
      ctx: var RuntimeContext): TaskAction =
    discard params
    let backend = ctx.backend
    let logger = ctx.logger
    return proc() =
      backend.leftClick()
      if logger != nil:
        logger.debug("Left click issued")
  )

  registry.registerAction("capture_window_target", proc(params: Table[string,
      string], ctx: var RuntimeContext): TaskAction =
    let targetName = params.getOrDefault("target", "").strip()
    let persist = parseBoolOpt(params, "persist", true)

    let backend = ctx.backend
    let logger = ctx.logger
    var targets = ctx.windowTargets
    let statePathOpt = ctx.windowTargetStatePath

    return proc() =
      if targetName.len == 0:
        if logger != nil:
          logger.warn("capture_window_target requires a 'target' parameter")
        return

      var hwnd: WindowHandle
      try:
        hwnd = backend.getActiveWindow()
      except CatchableError as e:
        if logger != nil:
          logger.error("Failed to capture active window", [("error", e.msg)])
        return

      if hwnd == 0:
        if logger != nil:
          logger.warn("No active window detected while capturing target", [(
              "target", targetName)])
        return

      updateStoredHwnd(targets, targetName, hwnd, logger)

      if persist and statePathOpt.isSome:
        saveWindowTargetState(statePathOpt.get(), targets, logger)
  )

  registry.registerAction("center_active_window", proc(params: Table[string,
      string], ctx: var RuntimeContext): TaskAction =
    discard params
    let backend = ctx.backend
    let logger = ctx.logger
    return proc() =
      let hwnd = backend.getActiveWindow()
      if hwnd == 0:
        if logger != nil:
          logger.warn("No active window to center")
        return

      if backend.centerWindowOnPrimaryMonitor(hwnd):
        if logger != nil:
          logger.info("Centered active window", [("title",
              backend.getWindowTitle(hwnd))])
      else:
        if logger != nil:
          logger.error("Failed to center window", [("title",
              backend.getWindowTitle(hwnd))])
  )

  registry.registerAction("exit_loop", proc(params: Table[string, string],
      ctx: var RuntimeContext): TaskAction =
    discard params
    let backend = ctx.backend
    let logger = ctx.logger
    return proc() =
      if logger != nil:
        logger.info("Requesting message loop exit")
      backend.postQuit()
  )

