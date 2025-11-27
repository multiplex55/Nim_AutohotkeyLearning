import ./actions
import ../core/logging
import ../core/runtime_context

type
  Plugin* = ref object of RootObj
    name*: string
    description*: string

  PluginManager* = ref object
    plugins: seq[Plugin]
    logger: Logger

method install*(plugin: Plugin, registry: var ActionRegistry, ctx: var RuntimeContext) {.base.} = discard
method shutdown*(plugin: Plugin, ctx: RuntimeContext) {.base.} = discard

proc newPluginManager*(logger: Logger = nil): PluginManager =
  PluginManager(plugins: @[], logger: logger)

proc registerPlugin*(manager: PluginManager, plugin: Plugin, registry: var ActionRegistry, ctx: var RuntimeContext) =
  if plugin == nil:
    return
  manager.plugins.add(plugin)
  plugin.install(registry, ctx)
  if manager.logger != nil:
    manager.logger.info("Plugin installed", [("name", plugin.name)])

proc shutdownPlugins*(manager: PluginManager, ctx: RuntimeContext) =
  for plugin in manager.plugins:
    plugin.shutdown(ctx)
    if manager.logger != nil:
      manager.logger.debug("Plugin shut down", [("name", plugin.name)])
