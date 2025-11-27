## Find a UIA button by name and click it.
## Compile with:
##   nim c -r -d:release -p:src examples/uia_button_click.nim

when not defined(windows):
  {.fatal: "This example only runs on Windows.".}

import std/[options, strformat, times]
import features/uia/uia

let targetName = "OK"

let automation = initUia()
defer: automation.shutdown()

echo fmt"Looking for a button named '{targetName}' ..."
let button = automation.waitElement(tsDescendants, automation.nameAndControlType(targetName, UIA_ButtonControlTypeId), 3.seconds)
if button.isNil:
  echo fmt"Button '{targetName}' not found. Try focusing the target window first."
  quit(QuitFailure)

if button.hasPattern(UIA_InvokePatternId, "Invoke"):
  button.invoke()
  echo fmt"Invoked '{targetName}' button."
else:
  echo fmt"Button '{targetName}' does not support Invoke."
