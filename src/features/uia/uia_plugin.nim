import ./uia
import ../actions
import ../plugins
import ../../core/runtime_context

## Plugin that owns the lifetime of a UI Automation session.
type
  UiaPlugin* = ref object of Plugin
    uia*: Uia

proc newUiaPlugin*(): UiaPlugin =
  UiaPlugin(name: "uia", description: "Windows UI Automation helpers")

method install*(plugin: UiaPlugin, registry: var ActionRegistry, ctx: RuntimeContext) =
  discard registry
  plugin.uia = initUia()
  if ctx.logger != nil:
    ctx.logger.info("Initialized UIA session", [("name", plugin.name)])

method shutdown*(plugin: UiaPlugin, ctx: RuntimeContext) =
  discard ctx
  if plugin.uia != nil:
    plugin.uia.shutdown()
    plugin.uia = nil
