import std/tables

import ../actions
import ../../core/logging
import ../../core/runtime_context
import ../../core/platform_backend
import ../../core/scheduler
import ../plugins

type
  WindowsHelpers* = ref object of Plugin

proc newWindowsHelpers*(): WindowsHelpers =
  WindowsHelpers(name: "windows_helpers", description: "Windows clipboard and window helpers")

method install*(plugin: WindowsHelpers, registry: var ActionRegistry, ctx: var RuntimeContext) =
  discard ctx
  registry.registerAction("active_window_info", proc(params: Table[string, string], ctx: var RuntimeContext): TaskAction =
    discard params
    return proc() =
      let hwnd = ctx.backend.getActiveWindow()
      if hwnd == 0:
        if ctx.logger != nil:
          ctx.logger.warn("No active window detected")
        return
      if ctx.logger != nil:
        ctx.logger.info("Active window info", [("title", ctx.backend.getWindowTitle(hwnd)), ("details", ctx.backend.describeWindow(hwnd))])
  )

  registry.registerAction("snap_active_center", proc(params: Table[string, string], ctx: var RuntimeContext): TaskAction =
    discard params
    return proc() =
      let hwnd = ctx.backend.getActiveWindow()
      if hwnd == 0:
        if ctx.logger != nil:
          ctx.logger.warn("Cannot snap center; no active window")
        return
      if ctx.backend.centerWindowOnPrimaryMonitor(hwnd):
        if ctx.logger != nil:
          ctx.logger.info("Snapped active window to center", [("title", ctx.backend.getWindowTitle(hwnd))])
      else:
        if ctx.logger != nil:
          ctx.logger.error("Failed to snap active window", [("title", ctx.backend.getWindowTitle(hwnd))])
  )

  registry.registerAction("screen_info", proc(params: Table[string, string], ctx: var RuntimeContext): TaskAction =
    discard params
    return proc() =
      let (w, h) = ctx.backend.getPrimaryScreenSize()
      if ctx.logger != nil:
        ctx.logger.info("Screen info", [("width", $w), ("height", $h)])
  )

method shutdown*(plugin: WindowsHelpers, ctx: RuntimeContext) =
  if ctx.logger != nil:
    ctx.logger.debug("Windows helpers plugin shutdown", [("name", plugin.name)])
