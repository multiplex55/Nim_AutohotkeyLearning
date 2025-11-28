import std/[options, strutils, tables, times]

import ./uia
import ../plugins
import ../../core/[logging, runtime_context]

when defined(windows):
  import winim/lean
  import winim/inc/uiautomation

  import ../actions
  import ../../core/window_targets
  import ../../platform/windows/processes as winProcesses

  type
    UiaPlugin* = ref object of Plugin
      uia*: Uia

  proc newUiaPlugin*(): UiaPlugin =
    UiaPlugin(name: "uia", description: "Windows UI Automation helpers")

  method install*(plugin: UiaPlugin, registry: var ActionRegistry,
      ctx: var RuntimeContext) =
    # Initialize the UIA session for this plugin.
    plugin.uia = initUia()

    # Register any UIA-based actions here using `registry.registerAction(...)`
    # (click-button, dump-under-mouse, etc.)


else:
  import ../actions

  type
    UiaPlugin* = ref object of Plugin

  proc newUiaPlugin*(): UiaPlugin =
    UiaPlugin(name: "uia", description: "Windows UI Automation helpers (disabled)")

  method install*(plugin: UiaPlugin, registry: var ActionRegistry,
      ctx: var RuntimeContext) =
    # Initialize the UIA session for this plugin.
    plugin.uia = initUia()

    # Register any UIA-based actions here using `registry.registerAction(...)`
    # (click-button, dump-under-mouse, etc.)

