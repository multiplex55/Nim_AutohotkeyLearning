## win/win.nim
## Cross-platform window helpers with Windows backend and stubs elsewhere.

import std/[monotimes, options, strformat, times]

when defined(windows):
  import winim/lean
  import winim/winstr

  const titleBufferLen = 512

  type
    WindowHandle* = HWND
    WindowRect* = object
      x*, y*, width*, height*: int
    WindowInfo* = object
      handle*: WindowHandle
      title*: string
      rect*: WindowRect

    WindowPollingOptions* = object
      ##
      ## Controls how aggressively helper functions (like `winWait`) poll for
      ## windows. Use this to throttle repeated enumeration in tight loops and
      ## avoid unnecessary CPU usage.
      pollInterval*: Duration   ## Time to sleep between polls
      debounce*: Duration       ## Minimum spacing between polls

    EnumContext = object
      acc: ptr seq[WindowInfo]
      titleBuf: ptr array[titleBufferLen, WCHAR]

  const defaultWindowPolling* = WindowPollingOptions(
    pollInterval: 100.milliseconds,
    debounce: 0.milliseconds
  )

  proc readTitleWithBuffer(hwnd: WindowHandle; buf: var array[titleBufferLen, WCHAR]): string =
    ## Read the title of a window using a reusable buffer to avoid repeated
    ## allocations when enumerating many windows.
    if hwnd == 0:
      return ""

    let copied = GetWindowText(hwnd, cast[LPWSTR](addr buf[0]), titleBufferLen.int32)
    if copied <= 0:
      return ""

    if copied < titleBufferLen:
      buf[copied] = WCHAR(0)
    $cast[LPWSTR](addr buf[0])

  proc readTitle(hwnd: WindowHandle): string =
    var buf: array[titleBufferLen, WCHAR]
    result = readTitleWithBuffer(hwnd, buf)

  proc windowInfoWithBuffer(hwnd: WindowHandle; buf: var array[titleBufferLen, WCHAR]): Option[WindowInfo] =
    if hwnd == 0:
      return none(WindowInfo)
    var r: RECT
    if GetWindowRect(hwnd, addr r) == 0:
      return none(WindowInfo)
    some(WindowInfo(handle: hwnd, title: readTitleWithBuffer(hwnd, buf), rect: toRect(r)))

  proc toRect(r: RECT): WindowRect =
    WindowRect(
      x: int(r.left),
      y: int(r.top),
      width: int(r.right - r.left),
      height: int(r.bottom - r.top)
    )

  proc windowInfo*(hwnd: WindowHandle): Option[WindowInfo] =
    var buf: array[titleBufferLen, WCHAR]
    result = windowInfoWithBuffer(hwnd, buf)

  proc listWindows*(includeUntitled = false): seq[WindowInfo] =
    proc enumProc(hwnd: HWND, l: LPARAM): WINBOOL {.stdcall.} =
      if IsWindowVisible(hwnd) == 0:
        return 1

      let ctx = cast[ptr EnumContext](l)
      let title = readTitleWithBuffer(hwnd, ctx.titleBuf[])
      if not includeUntitled and title.len == 0:
        return 1

      var r: RECT
      if GetWindowRect(hwnd, addr r) != 0:
        ctx.acc[][email protected](WindowInfo(
          handle: hwnd,
          title: title,
          rect: toRect(r)
        ))
      1

    var acc: seq[WindowInfo]
    var titleBuf: array[titleBufferLen, WCHAR]
    var ctx = EnumContext(acc: addr acc, titleBuf: addr titleBuf)
    discard EnumWindows(enumProc, cast[LPARAM](addr ctx))
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

  proc winWait*(title: string; timeout: Duration = 3.seconds;
                options: WindowPollingOptions = defaultWindowPolling): Option[WindowInfo] =
    let deadline = now() + timeout
    var lastPoll = getMonoTime() - options.debounce
    while now() < deadline:
      let nowTick = getMonoTime()
      let sinceLast = nowTick - lastPoll
      if options.debounce > 0.seconds and sinceLast < options.debounce:
        let sleepFor = (options.debounce - sinceLast).inMilliseconds
        if sleepFor > 0:
          sleep(int(sleepFor))
        continue

      lastPoll = nowTick
      let found = findWindowByTitle(title)
      if found.isSome:
        return found

      if options.pollInterval > 0.seconds:
        sleep(int(options.pollInterval.inMilliseconds))
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
    WindowPollingOptions* = object
      pollInterval*: Duration
      debounce*: Duration

  const defaultWindowPolling* = WindowPollingOptions(
    pollInterval: 0.seconds,
    debounce: 0.seconds
  )

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

  proc winWait*(title: string; timeout: Duration = 3.seconds;
                options: WindowPollingOptions = defaultWindowPolling): Option[WindowInfo] =
    discard title; discard timeout; discard options
    none(WindowInfo)

  proc windowInfo*(hwnd: WindowHandle): Option[WindowInfo] =
    discard hwnd
    none(WindowInfo)
