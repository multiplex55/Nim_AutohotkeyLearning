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
  registry.registerAction("active_window_info", proc(params: Table[string, string], ctx: var RuntimeContext): TaskAction =
    discard params
    let backend = ctx.backend
    let logger = ctx.logger
    return proc() =
      let hwnd = backend.getActiveWindow()
      if hwnd == 0:
        if logger != nil:
          logger.warn("No active window detected")
        return
      if logger != nil:
        logger.info("Active window info", [("title", backend.getWindowTitle(hwnd)), ("details", backend.describeWindow(hwnd))])
  )

  registry.registerAction("snap_active_center", proc(params: Table[string, string], ctx: var RuntimeContext): TaskAction =
    discard params
    let backend = ctx.backend
    let logger = ctx.logger
    return proc() =
      let hwnd = backend.getActiveWindow()
      if hwnd == 0:
        if logger != nil:
          logger.warn("Cannot snap center; no active window")
        return
      if backend.centerWindowOnPrimaryMonitor(hwnd):
        if logger != nil:
          logger.info("Snapped active window to center", [("title", backend.getWindowTitle(hwnd))])
      else:
        if logger != nil:
          logger.error("Failed to snap active window", [("title", backend.getWindowTitle(hwnd))])
  )

  registry.registerAction("screen_info", proc(params: Table[string, string], ctx: var RuntimeContext): TaskAction =
    discard params
    let backend = ctx.backend
    let logger = ctx.logger
    return proc() =
      let (w, h) = backend.getPrimaryScreenSize()
      if logger != nil:
        logger.info("Screen info", [("width", $w), ("height", $h)])
  )

method shutdown*(plugin: WindowsHelpers, ctx: RuntimeContext) =
  if ctx.logger != nil:
    ctx.logger.debug("Windows helpers plugin shutdown", [("name", plugin.name)])
