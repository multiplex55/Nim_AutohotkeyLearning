import ./logging
import ./scheduler
import ./platform_backend

type
  RuntimeContext* = object
    logger*: Logger
    scheduler*: Scheduler
    backend*: PlatformBackend
