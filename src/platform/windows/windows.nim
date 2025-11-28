## windows.nim
## AHK-like helpers for working with top-level windows on Windows.
##
## Features:
##   - Get the currently active (foreground) window
##   - Read a window's title
##   - Query a window's position/size
##   - Move/center a window
##   - Bring a window to the foreground
##   - Find a window by its exact title
##   - Describe a window (handle + title + rect)
##
## This module uses Win32 APIs via winim and is Windows-only.

when system.hostOS != "windows":
  {.error: "windows.nim only supports Windows.".}

import std/strformat
import winim/lean
import winim/winstr # for wstring helpers (T, nullTerminate, $)

type
  ## Opaque handle to a window (HWND).
  WindowHandle* = HWND

  ## Simple rectangle type for window geometry (screen coordinates).
  WindowRect* = object
    x*, y*: int          ## Top-left corner of the window
    width*, height*: int ## Size of the window


# ─────────────────────────────────────────────────────────────────────────────
# Basic queries
# ─────────────────────────────────────────────────────────────────────────────

proc getActiveWindow*(): WindowHandle =
  ## Returns the handle of the current foreground (active) window.
  ## Returns 0 if there is no active window.
  GetForegroundWindow()

proc getWindowTitle*(hwnd: WindowHandle): string =
  ## Returns the title text (caption) of the given top-level window.
  ##
  ## If the handle is 0 or the title cannot be read, an empty string is returned.
  if hwnd == 0 or IsWindow(hwnd) == 0:
    return ""

  ## Allocate a wide-char buffer using winim/winstr's T() template.
  ## T(512) gives us a wstring buffer of length 512 WCHARs.
  var buf = T(512)

  ## IMPORTANT: nMaxCount must be int32, so cast len(buf).
  let copied = GetWindowText(hwnd, &buf, int32(len(buf)))
  if copied <= 0:
    return ""

  ## Fix the internal length by scanning for the null terminator.
  nullTerminate(buf)

  ## Convert wstring -> Nim string (utf-8).
  result = $buf

proc getWindowRect*(hwnd: WindowHandle): WindowRect =
  ## Returns the bounding rectangle of the given window in screen coordinates.
  ##
  ## If the handle is invalid or the call fails, all fields are 0.
  var r: RECT

  if hwnd == 0 or IsWindow(hwnd) == 0 or GetWindowRect(hwnd, addr r) == 0:
    return WindowRect(x: 0, y: 0, width: 0, height: 0)

  WindowRect(
    x: int(r.left),
    y: int(r.top),
    width: int(r.right - r.left),
    height: int(r.bottom - r.top)
  )


# ─────────────────────────────────────────────────────────────────────────────
# Moving, centering, focusing
# ─────────────────────────────────────────────────────────────────────────────

proc moveWindow*(
    hwnd: WindowHandle;
    x, y, width, height: int;
    repaint: bool = true
  ): bool =
  ## Moves and resizes the window to the given rectangle.
  ##
  ## Returns true on success, false on failure (or hwnd = 0).
  if hwnd == 0 or IsWindow(hwnd) == 0:
    return false

  result = MoveWindow(
    hwnd,
    x.int32,
    y.int32,
    width.int32,
    height.int32,
    repaint
  ) != 0

proc centerWindowOnPrimaryMonitor*(hwnd: WindowHandle): bool =
  ## Centers the window on the primary monitor while keeping its current size.
  ##
  ## Returns true on success, false if hwnd is invalid or geometry cannot be read.
  if hwnd == 0 or IsWindow(hwnd) == 0:
    return false

  let r = getWindowRect(hwnd)
  if r.width <= 0 or r.height <= 0:
    return false

  let screenW = GetSystemMetrics(SM_CXSCREEN)
  let screenH = GetSystemMetrics(SM_CYSCREEN)

  let newX = (screenW - r.width) div 2
  let newY = (screenH - r.height) div 2

  result = moveWindow(hwnd, newX, newY, r.width, r.height)

proc bringToFront*(hwnd: WindowHandle): bool =
  ## Brings the window to the foreground (activates it).
  ##
  ## Returns true on success, false on failure or if hwnd is 0.
  if hwnd == 0 or IsWindow(hwnd) == 0:
    return false

  result = SetForegroundWindow(hwnd) != 0


# ─────────────────────────────────────────────────────────────────────────────
# Finding windows
# ─────────────────────────────────────────────────────────────────────────────

proc findWindowByTitleExact*(title: string): WindowHandle =
  ## Finds a top-level window whose title matches `title` exactly.
  ##
  ## Returns 0 if no matching window is found.
  ##
  ## NOTE:
  ##   This is a simple, exact match. If you want fuzzy/contains
  ##   matching, you can build that on top by enumerating windows.
  ##
  ## winim/winstr provides converters so we can pass `string`
  ## where LPWSTR is expected.
  FindWindowW(nil, title)

proc describeWindow*(hwnd: WindowHandle): string =
  ## Returns a human-readable summary of the window handle,
  ## including title and geometry.
  if hwnd == 0:
    return "HWND=0x0 (null window)"
  if IsWindow(hwnd) == 0:
    return &"HWND=0x{cast[uint](hwnd):x} (invalid window)"

  let title = getWindowTitle(hwnd)
  let r = getWindowRect(hwnd)

  result = &"HWND=0x{cast[uint](hwnd):x}, " &
           &"title=\"{title}\", x={r.x}, y={r.y}, w={r.width}, h={r.height}"
