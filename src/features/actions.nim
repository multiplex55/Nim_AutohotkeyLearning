import std/[strutils, tables]

import ../core/logging
import ../core/runtime_context
import ../core/platform_backend
import ../core/scheduler
import ../core/window_targets
import ./window_target_state

type
  ActionFactory* = proc(params: Table[string, string], ctx: var RuntimeContext): TaskAction

  ActionRegistry* = ref object
    factories: Table[string, ActionFactory]
    logger: Logger

proc newActionRegistry*(logger: Logger = nil): ActionRegistry =
  ActionRegistry(factories: initTable[string, ActionFactory](), logger: logger)

proc registerAction*(registry: ActionRegistry, name: string, factory: ActionFactory) =
  registry.factories[name.toLowerAscii()] = factory
  if registry.logger != nil:
    registry.logger.debug("Registered action", [("name", name)])

proc createAction*(registry: ActionRegistry, name: string, params: Table[string, string], ctx: var RuntimeContext): TaskAction =
  let key = name.toLowerAscii()
  if key notin registry.factories:
    if registry.logger != nil:
      registry.logger.warn("Unknown action requested", [("action", name)])
    return proc() = discard
  registry.factories[key](params, ctx)

proc parseIntOpt(params: Table[string, string], key: string, default: int = 0): int =
  if key in params:
    try:
      result = parseInt(params[key])
    except ValueError:
      result = default
  else:
    result = default

proc parseBoolOpt(params: Table[string, string], key: string, default: bool = true): bool =
  if key in params:
    let value = params[key].toLowerAscii()
    return value in ["1", "true", "yes", "on"]
  default

proc registerBuiltinActions*(registry: ActionRegistry) =
  ## Built-in actions that don't require plugins.
  registry.registerAction("start_process", proc(params: Table[string, string], ctx: var RuntimeContext): TaskAction =
    let cmd = params.getOrDefault("command", "")
    return proc() =
      if ctx.backend.startProcessDetached(cmd):
        if ctx.logger != nil:
          ctx.logger.info("Started process", [("command", cmd)])
      else:
        if ctx.logger != nil:
          ctx.logger.error("Failed to start process", [("command", cmd)])
  )

  registry.registerAction("kill_process", proc(params: Table[string, string], ctx: var RuntimeContext): TaskAction =
    let name = params.getOrDefault("name", "")
    return proc() =
      let killed = ctx.backend.killProcessesByName(name)
      if ctx.logger != nil:
        ctx.logger.info("Kill process result", [("name", name), ("killed", $killed)])
  )

  registry.registerAction("send_text", proc(params: Table[string, string], ctx: var RuntimeContext): TaskAction =
    let msg = params.getOrDefault("text", "")
    return proc() =
      ctx.backend.sendText(msg)
      if ctx.logger != nil:
        ctx.logger.info("Sent text", [("text", msg)])
  )

  registry.registerAction("move_mouse", proc(params: Table[string, string], ctx: var RuntimeContext): TaskAction =
    let x = parseIntOpt(params, "x")
    let y = parseIntOpt(params, "y")
    return proc() =
      if ctx.backend.setMousePos(x, y):
        if ctx.logger != nil:
          ctx.logger.info("Mouse moved", [("x", $x), ("y", $y)])
      else:
        if ctx.logger != nil:
          ctx.logger.error("Failed to move mouse", [("x", $x), ("y", $y)])
  )

  registry.registerAction("left_click", proc(params: Table[string, string], ctx: var RuntimeContext): TaskAction =
    discard params
    return proc() =
      ctx.backend.leftClick()
      if ctx.logger != nil:
        ctx.logger.debug("Left click issued")
  )

  registry.registerAction("capture_window_target", proc(params: Table[string, string], ctx: var RuntimeContext): TaskAction =
    let targetName = params.getOrDefault("target", "").strip()
    let persist = parseBoolOpt(params, "persist", true)

    return proc() =
      if targetName.len == 0:
        if ctx.logger != nil:
          ctx.logger.warn("capture_window_target requires a 'target' parameter")
        return

      var hwnd: WindowHandle
      try:
        hwnd = ctx.backend.getActiveWindow()
      except CatchableError as e:
        if ctx.logger != nil:
          ctx.logger.error("Failed to capture active window", [("error", e.msg)])
        return

      if hwnd == 0:
        if ctx.logger != nil:
          ctx.logger.warn("No active window detected while capturing target", [("target", targetName)])
        return

      updateStoredHwnd(ctx.windowTargets, targetName, hwnd, ctx.logger)

      if persist and ctx.windowTargetStatePath.isSome:
        saveWindowTargetState(ctx.windowTargetStatePath.get(), ctx.windowTargets, ctx.logger)
  )

  registry.registerAction("center_active_window", proc(params: Table[string, string], ctx: var RuntimeContext): TaskAction =
    discard params
    return proc() =
      let hwnd = ctx.backend.getActiveWindow()
      if hwnd == 0:
        if ctx.logger != nil:
          ctx.logger.warn("No active window to center")
        return
      if ctx.backend.centerWindowOnPrimaryMonitor(hwnd):
        if ctx.logger != nil:
          ctx.logger.info("Centered active window", [("title", ctx.backend.getWindowTitle(hwnd))])
      else:
        if ctx.logger != nil:
          ctx.logger.error("Failed to center window", [("title", ctx.backend.getWindowTitle(hwnd))])
  )

  registry.registerAction("exit_loop", proc(params: Table[string, string], ctx: var RuntimeContext): TaskAction =
    discard params
    return proc() =
      if ctx.logger != nil:
        ctx.logger.info("Requesting message loop exit")
      ctx.backend.postQuit()
  )
