## input/input.nim
## Cross-platform mouse and keyboard helpers with Windows SendInput backend.

import std/[options, sequtils, strutils, times]

when defined(windows):
  import winim/lean

  type
    MousePoint* = object
      x*, y*: int

    InputDelays* = object
      betweenEvents*: Duration
      betweenChars*: Duration

  const defaultDelays* = InputDelays(
    betweenEvents: 5.milliseconds,
    betweenChars: 20.milliseconds
  )

  proc sleepDelay(d: Duration) =
    if d > 0.seconds:
      sleep(int(d.inMilliseconds))

  proc normalizeCoordinate(value, max: int): int32 =
    ((value * 65535) div max).int32

  proc sendMouseInput(dx, dy: int; data: int32; flags: DWORD; absolute: bool) =
    var inp: INPUT
    inp.type = INPUT_MOUSE
    inp.mi.dwFlags = flags
    inp.mi.mouseData = data
    if absolute:
      let screenW = GetSystemMetrics(SM_CXSCREEN) - 1
      let screenH = GetSystemMetrics(SM_CYSCREEN) - 1
      inp.mi.dx = normalizeCoordinate(dx, screenW)
      inp.mi.dy = normalizeCoordinate(dy, screenH)
      inp.mi.dwFlags = inp.mi.dwFlags or MOUSEEVENTF_ABSOLUTE
    else:
      inp.mi.dx = dx.int32
      inp.mi.dy = dy.int32
    discard SendInput(1, addr inp, sizeof(INPUT).int32)

  proc moveMouse*(pos: MousePoint; relative = false;
      delays: InputDelays = defaultDelays) =
    if relative:
      sendMouseInput(pos.x, pos.y, 0'i32, MOUSEEVENTF_MOVE, false)
    else:
      sendMouseInput(pos.x, pos.y, 0'i32, MOUSEEVENTF_MOVE, true)
    sleepDelay(delays.betweenEvents)

  proc clickMouse*(button: string = "left"; pos: Option[MousePoint] = none(
      MousePoint); delays: InputDelays = defaultDelays) =
    if pos.isSome:
      moveMouse(pos.get, relative = false, delays = delays)
    var downFlag, upFlag: DWORD
    case button.toLowerAscii()
    of "left":
      downFlag = MOUSEEVENTF_LEFTDOWN
      upFlag = MOUSEEVENTF_LEFTUP
    of "right":
      downFlag = MOUSEEVENTF_RIGHTDOWN
      upFlag = MOUSEEVENTF_RIGHTUP
    of "middle":
      downFlag = MOUSEEVENTF_MIDDLEDOWN
      upFlag = MOUSEEVENTF_MIDDLEUP
    else:
      downFlag = MOUSEEVENTF_LEFTDOWN
      upFlag = MOUSEEVENTF_LEFTUP
    sendMouseInput(0, 0, 0'i32, downFlag, false)
    sleepDelay(delays.betweenEvents)
    sendMouseInput(0, 0, 0'i32, upFlag, false)
    sleepDelay(delays.betweenEvents)

  proc dragMouse*(startPos, endPos: MousePoint; steps = 10;
      delays: InputDelays = defaultDelays) =
    moveMouse(startPos, relative = false, delays = delays)
    sendMouseInput(0, 0, 0'i32, MOUSEEVENTF_LEFTDOWN, false)
    if steps < 1:
      steps = 1
    let dx = (endPos.x - startPos.x) div steps
    let dy = (endPos.y - startPos.y) div steps
    var current = startPos
    for _ in 0 ..< steps:
      current.x += dx
      current.y += dy
      moveMouse(current, relative = false, delays = delays)
    moveMouse(endPos, relative = false, delays = delays)
    sendMouseInput(0, 0, 0'i32, MOUSEEVENTF_LEFTUP, false)
    sleepDelay(delays.betweenEvents)

  proc scrollMouse*(deltaY: int; deltaX: int = 0;
      delays: InputDelays = defaultDelays) =
    if deltaY != 0:
      sendMouseInput(0, 0, deltaY.int32, MOUSEEVENTF_WHEEL, false)
      sleepDelay(delays.betweenEvents)
    if deltaX != 0:
      sendMouseInput(0, 0, deltaX.int32, MOUSEEVENTF_HWHEEL, false)
      sleepDelay(delays.betweenEvents)

  proc sendKeyEvent(vk: int; keyUp: bool; delays: InputDelays) =
    var inp: INPUT
    inp.type = INPUT_KEYBOARD
    inp.ki.wVk = WORD(vk)
    if keyUp:
      inp.ki.dwFlags = KEYEVENTF_KEYUP
    discard SendInput(1, addr inp, sizeof(INPUT).int32)
    sleepDelay(delays.betweenEvents)

  proc sendKeys*(keys: openArray[int]; delays: InputDelays = defaultDelays) =
    for key in keys:
      sendKeyEvent(key, false, delays)
    for key in keys.reversed:
      sendKeyEvent(key, true, delays)

  proc hotkey*(keys: openArray[int]; delays: InputDelays = defaultDelays) =
    sendKeys(keys, delays)

  proc sendChar(ch: char; delays: InputDelays) =
    let scan = VkKeyScanW(WCHAR(ch))
    if scan == -1:
      return
    let vk = scan and 0xff
    let shiftRequired = (scan and 0x0100) != 0
    if shiftRequired:
      sendKeyEvent(VK_SHIFT, false, delays)
    sendKeyEvent(vk, false, delays)
    sendKeyEvent(vk, true, delays)
    if shiftRequired:
      sendKeyEvent(VK_SHIFT, true, delays)

  proc typeText*(text: string; delays: InputDelays = defaultDelays) =
    for ch in text:
      sendChar(ch, delays)
      sleepDelay(delays.betweenChars)

else:
  type
    MousePoint* = object
      x*, y*: int
    InputDelays* = object
      betweenEvents*: Duration
      betweenChars*: Duration

  const defaultDelays* = InputDelays(betweenEvents: 0.seconds,
      betweenChars: 0.seconds)
  const unsupported = "Mouse/keyboard helpers are only implemented for Windows targets."

  proc unsupportedProc[T](name: string): T =
    raise newException(OSError, &"{unsupported} Tried to call '{name}'.")

  proc moveMouse*(pos: MousePoint; relative = false;
      delays: InputDelays = defaultDelays) =
    discard pos; discard relative; discard delays
    unsupportedProc[void]("moveMouse")

  proc clickMouse*(button: string = "left"; pos: Option[MousePoint] = none(
      MousePoint); delays: InputDelays = defaultDelays) =
    discard button; discard pos; discard delays
    unsupportedProc[void]("clickMouse")

  proc dragMouse*(startPos, endPos: MousePoint; steps = 10;
      delays: InputDelays = defaultDelays) =
    discard startPos; discard endPos; discard steps; discard delays
    unsupportedProc[void]("dragMouse")

  proc scrollMouse*(deltaY: int; deltaX: int = 0;
      delays: InputDelays = defaultDelays) =
    discard deltaY; discard deltaX; discard delays
    unsupportedProc[void]("scrollMouse")

  proc sendKeys*(keys: openArray[int]; delays: InputDelays = defaultDelays) =
    discard keys; discard delays
    unsupportedProc[void]("sendKeys")

  proc hotkey*(keys: openArray[int]; delays: InputDelays = defaultDelays) =
    discard keys; discard delays
    unsupportedProc[void]("hotkey")

  proc typeText*(text: string; delays: InputDelays = defaultDelays) =
    discard text; discard delays
    unsupportedProc[void]("typeText")
