import std/[json, strformat, strutils, times]

type
  LogLevel* = enum
    llTrace = 0, llDebug, llInfo, llWarn, llError

  Logger* = ref object
    level*: LogLevel
    structured*: bool

proc newLogger*(level = llInfo, structured = false): Logger =
  ## Create a new logger instance with the desired verbosity and output mode.
  Logger(level: level, structured: structured)

proc setLogLevel*(logger: Logger, level: LogLevel) =
  ## Update the logger's level directly from a LogLevel enum.
  if logger == nil:
    return
  logger.level = level

proc setLogLevel*(logger: Logger, levelName: string) =
  ## Update the logger's level from a case-insensitive name.
  if logger == nil:
    return
  let lowered = levelName.toLowerAscii()
  case lowered
  of "trace": logger.level = llTrace
  of "debug": logger.level = llDebug
  of "info": logger.level = llInfo
  of "warn", "warning": logger.level = llWarn
  of "error", "err": logger.level = llError
  else: discard

proc shouldLog(logger: Logger, level: LogLevel): bool =
  level >= logger.level


proc nowIso(): string =
  getTime().utc.format("yyyy-MM-dd'T'HH:mm:ss'.'fff'Z'")

proc log*(logger: Logger, level: LogLevel, message: string, fields: openArray[(
    string, string)] = []) =
  ## Emit a log line honoring the configured level and structure.
  if logger == nil or not logger.shouldLog(level):
    return

  if logger.structured:
    var payload = %*{
      "ts": nowIso(),
      "level": $level,
      "msg": message
    }
    if fields.len > 0:
      var extra: JsonNode = newJObject()
      for (k, v) in fields:
        extra[k] = %v
      payload["fields"] = extra
    echo payload
  else:
    var extras = ""
    if fields.len > 0:
      var parts: seq[string] = @[]
      for (k, v) in fields:
        parts.add(&"{k}={v}")
      extras = " [" & parts.join(", ") & "]"
    echo &"[{nowIso()}] {level}: {message}{extras}"

proc trace*(logger: Logger, message: string, fields: openArray[(string,
    string)] = []) =
  logger.log(llTrace, message, fields)

proc debug*(logger: Logger, message: string, fields: openArray[(string,
    string)] = []) =
  logger.log(llDebug, message, fields)

proc info*(logger: Logger, message: string, fields: openArray[(string,
    string)] = []) =
  logger.log(llInfo, message, fields)

proc warn*(logger: Logger, message: string, fields: openArray[(string,
    string)] = []) =
  logger.log(llWarn, message, fields)

proc error*(logger: Logger, message: string, fields: openArray[(string,
    string)] = []) =
  logger.log(llError, message, fields)
