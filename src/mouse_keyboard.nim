## mouse_keyboard.nim
## AHK-like mouse and keyboard helpers for Windows.
##
## Features:
##   - Query and set mouse position
##   - Left/right click (at current position or at a specific point)
##   - Basic key down/up / key press
##   - Simple ASCII text sending (letters, digits, space, newline)
##
## This is intentionally small and focused. It uses the older
## `mouse_event` and `keybd_event` APIs which are easy to call
## and perfectly fine for automation purposes.

when system.hostOS != "windows":
  {.error: "mouse_keyboard.nim only supports Windows.".}

import winim/lean

type
  ## Simple struct to hold mouse coordinates in screen space.
  MousePoint* = object
    x*: int
    y*: int

# ─────────────────────────────────────────────────────────────────────────────
# Raw WinAPI wrappers (mouse_event / keybd_event)
# ─────────────────────────────────────────────────────────────────────────────

## Wrapper around the Win32 mouse_event function.
## We give it a Nim-friendly name and signature.
proc mouseEvent(dwFlags, dx, dy, dwData: DWORD; dwExtraInfo: ULONG_PTR)
  {.stdcall, dynlib: "user32", importc: "mouse_event".}

## Wrapper around the Win32 keybd_event function.
proc keybdEvent(bVk: BYTE; bScan: BYTE; dwFlags: DWORD; dwExtraInfo: ULONG_PTR)
  {.stdcall, dynlib: "user32", importc: "keybd_event".}

## Mouse event flags (subset).
const
  MOUSEEVENTF_LEFTDOWN*  = DWORD(0x0002)
  MOUSEEVENTF_LEFTUP*    = DWORD(0x0004)
  MOUSEEVENTF_RIGHTDOWN* = DWORD(0x0008)
  MOUSEEVENTF_RIGHTUP*   = DWORD(0x0010)

## Keyboard event flags (subset).
const
  KEYEVENTF_KEYUP* = DWORD(0x0002)


# ─────────────────────────────────────────────────────────────────────────────
# Mouse helpers
# ─────────────────────────────────────────────────────────────────────────────

proc getMousePos*(): MousePoint =
  ## Returns the current mouse position in screen coordinates.
  ##
  ## If GetCursorPos fails (rare), returns (0, 0).
  var pt: POINT
  if GetCursorPos(addr pt) != 0:
    result = MousePoint(x: int(pt.x), y: int(pt.y))
  else:
    result = MousePoint(x: 0, y: 0)


proc setMousePos*(x, y: int): bool =
  ## Moves the mouse cursor to the absolute screen coordinates (x, y).
  ##
  ## Returns true on success, false otherwise.
  result = SetCursorPos(int32(x), int32(y)) != 0


proc leftClick*() =
  ## Performs a left mouse click at the current cursor position.
  mouseEvent(
    MOUSEEVENTF_LEFTDOWN or MOUSEEVENTF_LEFTUP,
    DWORD(0), DWORD(0), DWORD(0), ULONG_PTR(0)
  )

proc rightClick*() =
  ## Performs a right mouse click at the current cursor position.
  mouseEvent(
    MOUSEEVENTF_RIGHTDOWN or MOUSEEVENTF_RIGHTUP,
    DWORD(0), DWORD(0), DWORD(0), ULONG_PTR(0)
  )

proc leftClickAt*(x, y: int) =
  ## Moves the cursor to (x, y) and performs a left click.
  discard setMousePos(x, y)
  leftClick()

proc rightClickAt*(x, y: int) =
  ## Moves the cursor to (x, y) and performs a right click.
  discard setMousePos(x, y)
  rightClick()


# ─────────────────────────────────────────────────────────────────────────────
# Keyboard helpers
# ─────────────────────────────────────────────────────────────────────────────

proc keyDown*(vk: int) =
  ## Simulate key down for the given virtual-key code.
  ##
  ## Example:
  ##   keyDown(VK_SHIFT)
  keybdEvent(BYTE(vk), BYTE(0), DWORD(0), ULONG_PTR(0))

proc keyUp*(vk: int) =
  ## Simulate key up for the given virtual-key code.
  ##
  ## Example:
  ##   keyUp(VK_SHIFT)
  keybdEvent(BYTE(vk), BYTE(0), KEYEVENTF_KEYUP, ULONG_PTR(0))

proc sendKeyPress*(vk: int) =
  ## Send a full key press (down + up) for the given virtual-key code.
  ##
  ## Example:
  ##   sendKeyPress(VK_RETURN)
  keyDown(vk)
  keyUp(vk)


proc sendText*(text: string) =
  ## Very simple text sender for US-layout-ish keyboards.
  ##
  ## Supports:
  ##   - 'a'..'z'   -> sent as lowercase letters
  ##   - 'A'..'Z'   -> sent as uppercase letters (Shift + key)
  ##   - '0'..'9'   -> sent as digits
  ##   - ' '        -> space
  ##   - '\n'       -> Enter
  ##
  ## Any other characters are currently ignored.

  # Helper: map a letter (a-z or A-Z) to its virtual-key code (VK_A..VK_Z).
  proc letterVk(ch: char): int =
    ## On Windows, VK codes for letters are always the uppercase code.
    ## 'a'..'z' map to 'A'..'Z' by subtracting 32 in ASCII.
    if ch in 'a'..'z':
      result = ord(ch) - 32     # 'a'(97) -> 'A'(65)
    elif ch in 'A'..'Z':
      result = ord(ch)
    else:
      result = ord(ch)

  for ch in text:
    case ch
    of 'a'..'z':
      # Lowercase letter: press the key without Shift.
      let vk = letterVk(ch)
      sendKeyPress(vk)

    of 'A'..'Z':
      # Uppercase letter: hold Shift while pressing the key.
      let vk = letterVk(ch)
      keyDown(VK_SHIFT)
      sendKeyPress(vk)
      keyUp(VK_SHIFT)

    of '0'..'9':
      # Digits map directly to their VK_* codes on US layout.
      sendKeyPress(ord(ch))

    of ' ':
      sendKeyPress(VK_SPACE)

    of '\n':
      sendKeyPress(VK_RETURN)

    else:
      # Unsupported character – ignore for now.
      discard

