import std/tables
import std/options

import ./logging
import ./scheduler
import ./platform_backend
import ./window_targets

type
  RuntimeContext* = object
    logger*: Logger
    scheduler*: Scheduler
    backend*: PlatformBackend
    windowTargets*: Table[string, WindowTarget]
    windowTargetStatePath*: Option[string]
