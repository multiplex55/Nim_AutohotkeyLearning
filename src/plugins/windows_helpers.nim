import winim/lean
import std/tables

import ../actions
import ../logging
import ../mouse_keyboard
import ../windows
import ../plugins

type
  WindowsHelpers* = ref object of Plugin

proc newWindowsHelpers*(): WindowsHelpers =
  WindowsHelpers(name: "windows_helpers", description: "Windows clipboard and window helpers")

method install*(plugin: WindowsHelpers, registry: var ActionRegistry, ctx: RuntimeContext) =
  discard ctx
  registry.registerAction("active_window_info", proc(params: Table[string, string], ctx: RuntimeContext): TaskAction =
    discard params
    return proc() =
      let hwnd = getActiveWindow()
      if hwnd == 0:
        if ctx.logger != nil:
          ctx.logger.warn("No active window detected")
        return
      if ctx.logger != nil:
        ctx.logger.info("Active window info", [("title", getWindowTitle(hwnd)), ("details", describeWindow(hwnd))])
  )

  registry.registerAction("snap_active_center", proc(params: Table[string, string], ctx: RuntimeContext): TaskAction =
    discard params
    return proc() =
      let hwnd = getActiveWindow()
      if hwnd == 0:
        if ctx.logger != nil:
          ctx.logger.warn("Cannot snap center; no active window")
        return
      if centerWindowOnPrimaryMonitor(hwnd):
        if ctx.logger != nil:
          ctx.logger.info("Snapped active window to center", [("title", getWindowTitle(hwnd))])
      else:
        if ctx.logger != nil:
          ctx.logger.error("Failed to snap active window", [("title", getWindowTitle(hwnd))])
  )

  registry.registerAction("screen_info", proc(params: Table[string, string], ctx: RuntimeContext): TaskAction =
    discard params
    return proc() =
      let w = GetSystemMetrics(SM_CXSCREEN)
      let h = GetSystemMetrics(SM_CYSCREEN)
      if ctx.logger != nil:
        ctx.logger.info("Screen info", [("width", $w), ("height", $h)])
  )

method shutdown*(plugin: WindowsHelpers, ctx: RuntimeContext) =
  if ctx.logger != nil:
    ctx.logger.debug("Windows helpers plugin shutdown", [("name", plugin.name)])
