## Demonstrates a small mouse macro: move, click, and drag.
## Compile with:
##   nim c -r -d:release -p:src examples/mouse_macro.nim

when not defined(windows):
  {.fatal: "This example only runs on Windows.".}

import std/[strformat, times]
import platform/windows/mouse_keyboard

proc pause(ms: int) = sleep(ms)

let center = MousePoint(x: 500, y: 400)
let dragTarget = MousePoint(x: 800, y: 600)

echo "Moving mouse to center and clicking..."
discard setMousePos(center.x, center.y)
leftClick()
pause(500)

echo fmt"Dragging from ({center.x}, {center.y}) to ({dragTarget.x}, {dragTarget.y})"
discard setMousePos(center.x, center.y)
# Simple stepped drag
for step in 0 .. 10:
  let lerpX = center.x + ((dragTarget.x - center.x) * step) div 10
  let lerpY = center.y + ((dragTarget.y - center.y) * step) div 10
  discard setMousePos(lerpX, lerpY)
  pause(30)
leftClick()

echo "Mouse macro complete."
