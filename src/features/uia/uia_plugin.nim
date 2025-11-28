import ../actions
import ../plugins
import ../../core/runtime_context
import ./uia

when defined(windows):
  # Add Windows-specific UIA helpers here when needed.

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

