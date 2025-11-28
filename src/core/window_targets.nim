import std/[options, tables]

import ./logging

## Represents a logical window target that can be resolved at runtime.
type
  WindowTarget* = object
    name*: string
    title*: Option[string]
    titleContains*: Option[string]
    className*: Option[string]
    processName*: Option[string]
    storedHwnd*: Option[int]

proc newWindowTarget*(name: string): WindowTarget =
  ## Create an empty window target with only a name populated.
  WindowTarget(
    name: name,
    title: none(string),
    titleContains: none(string),
    className: none(string),
    processName: none(string),
    storedHwnd: none(int)
  )

proc updateStoredHwnd*(targets: var Table[string, WindowTarget], name: string,
    hwnd: int, logger: Logger) =
  ## Update or insert a window target with a stored HWND value.
  var target =
    if name in targets:
      targets[name]
    else:
      newWindowTarget(name)

  target.storedHwnd = some(hwnd)
  targets[name] = target

  if logger != nil:
    logger.info("Captured window target", [("name", name), ("hwnd", $hwnd)])
