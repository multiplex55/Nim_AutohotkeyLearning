import std/[options, parseutils, strformat, strutils]

type
  CliOptions* = object
    uiaDemo*: bool
    uiaMaxDepth*: int
    configPath*: Option[string]

proc parseCliArgs*(args: seq[string]): CliOptions =
  ## Parse command-line arguments for the autohotkey runner.
  ##
  ## Recognized flags:
  ## * --uia-demo: runs the UIA demo instead of loading a config file.
  ## * --uia-max-depth <int> or --uia-max-depth=<int>: overrides demo traversal depth.
  ##
  ## The first non-flag argument is treated as the config path.
  result.uiaMaxDepth = 4

  var idx = 0
  while idx < args.len:
    let arg = args[idx]
    case arg
    of "--uia-demo":
      result.uiaDemo = true
    of "--uia-max-depth":
      if idx + 1 >= args.len:
        raise newException(ValueError, "--uia-max-depth requires an integer argument")
      inc idx
      let val = args[idx]
      if parseInt(val, result.uiaMaxDepth) == 0:
        raise newException(ValueError, fmt"Invalid max depth: {val}")
    else:
      if arg.startsWith("--uia-max-depth="):
        let parts = arg.split('=')
        if parts.len < 2 or parts[1].len == 0 or parseInt(parts[1], result.uiaMaxDepth) == 0:
          raise newException(ValueError, fmt"Invalid max depth: {arg}")
      elif arg.startsWith("--"):
        raise newException(ValueError, fmt"Unrecognized flag: {arg}")
      elif result.configPath.isSome:
        raise newException(ValueError, "Only one config path is supported")
      else:
        result.configPath = some(arg)
    inc idx
