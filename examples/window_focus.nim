## Bring a window with a specific title to the foreground and type into it.
## Compile with:
##   nim c -r -d:release -p:src examples/window_focus.nim

when not defined(windows):
  {.fatal: "This example only runs on Windows.".}

import std/[strformat]
import platform/windows/windows as winWindows
import platform/windows/mouse_keyboard

let targetTitle = "Untitled - Notepad"

let hwnd = winWindows.findWindowByTitleExact(targetTitle)
if hwnd == 0:
  echo fmt"Could not find a window titled '{targetTitle}'. Start Notepad and try again."
  quit(QuitFailure)

if winWindows.bringToFront(hwnd):
  echo fmt"Brought '{targetTitle}' to the foreground."
  sendText("Focused by Nim AHK Toolkit!\n")
else:
  echo fmt"Failed to activate '{targetTitle}'."
  quit(QuitFailure)
