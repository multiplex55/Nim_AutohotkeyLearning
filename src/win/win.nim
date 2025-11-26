## win/win.nim
## Cross-platform window helpers with Windows backend and stubs elsewhere.

import std/[options, strformat, times]

when defined(windows):
  import winim/lean
  import winim/winstr

  type
    WindowHandle* = HWND
    WindowRect* = object
      x*, y*, width*, height*: int
    WindowInfo* = object
      handle*: WindowHandle
      title*: string
      rect*: WindowRect

  proc toRect(r: RECT): WindowRect =
    WindowRect(
      x: int(r.left),
      y: int(r.top),
      width: int(r.right - r.left),
      height: int(r.bottom - r.top)
    )

  proc readTitle(hwnd: WindowHandle): string =
    var buf = T(512)
    let copied = GetWindowText(hwnd, &buf, int32(len(buf)))
    if copied <= 0:
      return ""
    nullTerminate(buf)
    $buf

  proc windowInfo*(hwnd: WindowHandle): Option[WindowInfo] =
    if hwnd == 0:
      return none(WindowInfo)
    var r: RECT
    if GetWindowRect(hwnd, addr r) == 0:
      return none(WindowInfo)
    some(WindowInfo(handle: hwnd, title: readTitle(hwnd), rect: toRect(r)))

  proc listWindows*(includeUntitled = false): seq[WindowInfo] =
    proc enumProc(hwnd: HWND, l: LPARAM): WINBOOL {.stdcall.} =
      if IsWindowVisible(hwnd) == 0:
        return 1
      let title = readTitle(hwnd)
      if not includeUntitled and title.len == 0:
        return 1
      var r: RECT
      if GetWindowRect(hwnd, addr r) != 0:
        cast[ptr seq[WindowInfo]](l)[][email protected](WindowInfo(
          handle: hwnd,
          title: title,
          rect: toRect(r)
        ))
      1
    var acc: seq[WindowInfo]
    discard EnumWindows(enumProc, cast[LPARAM](addr acc))
    acc

  proc findWindowByTitle*(title: string): Option[WindowInfo] =
    for win in listWindows(includeUntitled = true):
      if win.title == title:
        return some(win)
    none(WindowInfo)

  proc findWindowByHandle*(hwnd: WindowHandle): Option[WindowInfo] =
    windowInfo(hwnd)

  proc activateWindow*(hwnd: WindowHandle): bool =
    if hwnd == 0:
      return false
    SetForegroundWindow(hwnd) != 0

  proc moveResizeWindow*(
    hwnd: WindowHandle,
    x, y, width, height: int,
    repaint = true
  ): bool =
    if hwnd == 0:
      return false
    MoveWindow(hwnd, x.int32, y.int32, width.int32, height.int32, repaint) != 0

  proc winWait*(title: string; timeout: Duration = 3.seconds): Option[WindowInfo] =
    let deadline = now() + timeout
    while now() < deadline:
      let found = findWindowByTitle(title)
      if found.isSome:
        return found
      sleep(100)
    none(WindowInfo)

else:
  type
    WindowHandle* = int
    WindowRect* = object
      x*, y*, width*, height*: int
    WindowInfo* = object
      handle*: WindowHandle
      title*: string
      rect*: WindowRect

  const unsupported = "Window operations are only implemented for Windows targets." 

  proc unsupportedProc[T](name: string): T =
    raise newException(OSError, &"{unsupported} Tried to call '{name}'.")

  proc listWindows*(includeUntitled = false): seq[WindowInfo] =
    discard includeUntitled
    @[]

  proc findWindowByTitle*(title: string): Option[WindowInfo] =
    discard title
    none(WindowInfo)

  proc findWindowByHandle*(hwnd: WindowHandle): Option[WindowInfo] =
    discard hwnd
    none(WindowInfo)

  proc activateWindow*(hwnd: WindowHandle): bool =
    discard hwnd
    unsupportedProc[bool]("activateWindow")

  proc moveResizeWindow*(
    hwnd: WindowHandle,
    x, y, width, height: int,
    repaint = true
  ): bool =
    discard hwnd; discard x; discard y; discard width; discard height; discard repaint
    unsupportedProc[bool]("moveResizeWindow")

  proc winWait*(title: string; timeout: Duration = 3.seconds): Option[WindowInfo] =
    discard title; discard timeout
    none(WindowInfo)

  proc windowInfo*(hwnd: WindowHandle): Option[WindowInfo] =
    discard hwnd
    none(WindowInfo)
