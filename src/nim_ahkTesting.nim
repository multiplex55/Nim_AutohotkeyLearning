import std/[os, osproc, strformat, strutils]

import ./main as app

when defined(windows):
  import winim/inc/uiautomationclient
  import ./platform/windows/windows as winWindows
  import ./features/uia/uia

proc usage() =
  echo "Nim AHK Toolkit CLI"
  echo "Usage: nim_ahkTesting <command> [args...]"
  echo ""
  echo "Commands:"
  echo "  run-config [path]   Run hotkeys from a TOML config (defaults to examples/hotkeys.toml)"
  echo "  run-example <path>  Compile + run a Nim example (e.g. examples/window_focus.nim)"
  echo "  list-windows        Print visible top-level windows (Windows only)"
  echo "  list-elements       Print top-level UIA elements (Windows only)"
  echo "  help                Show this help"

proc runConfig(path: string) =
  if not fileExists(path):
    echo fmt"Config file {path} not found."
    quit(QuitFailure)
  discard app.setupHotkeys(path)

proc runExample(path: string) =
  let normalized =
    if fileExists(path):
      path
    else:
      let withExt = if path.endsWith(".nim") : path else: path & ".nim"
      joinPath("examples", withExt)

  if not fileExists(normalized):
    echo fmt"Example {normalized} not found."
    quit(QuitFailure)

  let cmd = @["nim", "c", "-r", "-d:release", "-p:src", normalized]
  echo fmt"Running: {cmd.join(\" \" )}"
  discard execShellCmd(cmd.join(" "))
  quit(QuitSuccess)

when defined(windows):
  proc listWindows() =
    let windows = winWindows.enumerateWindows()
    if windows.len == 0:
      echo "No visible windows detected."
    for w in windows:
      echo fmt"0x{cast[uint](w.handle):x}  {w.title}"

  proc listElements() =
    var automation: Uia
    try:
      automation = initUia()
      let root = automation.rootElement()
      let children = automation.findAll(tsChildren, automation.trueCondition(), root)
      if children.len == 0:
        echo "No UIA children under the desktop root."
      for el in children:
        let name = el.currentName()
        let className = el.currentClassName()
        let controlType = el.currentControlType()
        echo fmt"{name}  (class={className}, controlType={controlType})"
    finally:
      if automation != nil:
        automation.shutdown()
else:
  proc listWindows() =
    echo "Window listing is only available on Windows."

  proc listElements() =
    echo "UI Automation listing is only available on Windows."

when isMainModule:
  if paramCount() == 0:
    usage()
    quit(QuitFailure)

  case paramStr(1).toLowerAscii()
  of "run-config":
    let path = if paramCount() >= 2: paramStr(2) else: app.DEFAULT_CONFIG
    runConfig(path)
  of "run-example":
    if paramCount() < 2:
      echo "Please provide an example path (e.g. examples/window_focus.nim)."
      quit(QuitFailure)
    runExample(paramStr(2))
  of "list-windows":
    listWindows()
  of "list-elements":
    listElements()
  of "help", "-h", "--help":
    usage()
  else:
    echo fmt"Unknown command: {paramStr(1)}"
    usage()
    quit(QuitFailure)
