import std/[options, strutils, tables, times]

import ./uia
import ../plugins
import ../../core/[logging, runtime_context]

when uiaHeadersAvailable:
  import winim/lean
  import winim/inc/uiautomationclient

  import ../actions
  import ../../core/window_targets
  import ../../platform/windows/processes as winProcesses

  type
    UiaPlugin* = ref object of Plugin
      uia*: Uia

  proc newUiaPlugin*(): UiaPlugin =
    UiaPlugin(name: "uia", description: "Windows UI Automation helpers")

  proc installUiaPlugin*(registry: var ActionRegistry; ctx: var RuntimeContext) =
    ## UIA support requires the winim UIAutomation headers. When present we simply initialize
    ## the automation object so downstream code can build specific actions as needed.
    let plugin = newUiaPlugin()
    plugin.uia = initUia()
    discard ctx.registerPlugin(plugin)

else:
  import ../actions

  type
    UiaPlugin* = ref object of Plugin

  proc newUiaPlugin*(): UiaPlugin =
    UiaPlugin(name: "uia", description: "Windows UI Automation helpers (disabled)")

  proc installUiaPlugin*(registry: var ActionRegistry; ctx: var RuntimeContext) =
    ## When UI Automation headers are unavailable, register a no-op plugin and warn once.
    discard ctx.registerPlugin(newUiaPlugin())
    if ctx.logger != nil:
      warn(ctx.logger, "UI Automation headers not found; UIA actions are disabled", [])
