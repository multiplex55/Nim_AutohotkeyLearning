when system.hostOS != "windows":
  {.error: "UIA highlight overlay is only supported on Windows.".}

import std/[options, strutils]

import winim/lean
import winim/inc/winuser

import ../../core/logging
import ../uia/uia

const
  overlayClassName = "UiaInspectorHighlightOverlay"
  overlayTimerId = 1'u
  overlayColorKey = COLORREF(0x00F1F2F3) ## Rare color used as transparency key

var overlayClassAtom: ATOM

proc overlayWndProc(hwnd: HWND; msg: UINT; wParam: WPARAM;
    lParam: LPARAM): LRESULT {.stdcall.}

proc ensureOverlayClass(logger: Logger) =
  if overlayClassAtom != 0:
    return

  var wc: WNDCLASSEXW
  wc.cbSize = UINT(sizeof(wc))
  wc.style = CS_HREDRAW or CS_VREDRAW
  wc.lpfnWndProc = cast[WNDPROC](overlayWndProc)
  wc.cbClsExtra = 0
  wc.cbWndExtra = 0
  wc.hInstance = GetModuleHandleW(nil)
  wc.hIcon = 0
  wc.hCursor = LoadCursorW(0, IDC_ARROW)
  wc.hbrBackground = cast[HBRUSH](GetStockObject(NULL_BRUSH))
  wc.lpszMenuName = nil
  wc.lpszClassName = overlayClassName

  overlayClassAtom = RegisterClassExW(addr wc)
  if overlayClassAtom == 0 and logger != nil:
    logger.error("Failed to register overlay window class")

proc overlayWndProc(hwnd: HWND; msg: UINT; wParam: WPARAM;
    lParam: LPARAM): LRESULT {.stdcall.} =
  case msg
  of WM_PAINT:
    var ps: PAINTSTRUCT
    let hdc = BeginPaint(hwnd, addr ps)
    var rect: RECT
    discard GetClientRect(hwnd, addr rect)

    let color = COLORREF(GetWindowLongPtrW(hwnd, GWLP_USERDATA))

    let background = CreateSolidBrush(overlayColorKey)
    discard FillRect(hdc, addr rect, background)
    discard DeleteObject(background)

    let pen = CreatePen(PS_SOLID, 3, color)
    let oldPen = SelectObject(hdc, pen)
    let oldBrush = SelectObject(hdc, GetStockObject(NULL_BRUSH))

    discard Rectangle(hdc, rect.left, rect.top, rect.right, rect.bottom)

    discard SelectObject(hdc, oldPen)
    discard SelectObject(hdc, oldBrush)
    discard DeleteObject(pen)
    discard EndPaint(hwnd, addr ps)
    return 0
  of WM_TIMER:
    KillTimer(hwnd, overlayTimerId)
    DestroyWindow(hwnd)
    return 0
  of WM_DESTROY:
    KillTimer(hwnd, overlayTimerId)
  else:
    discard

  result = DefWindowProcW(hwnd, msg, wParam, lParam)

proc parseColorRef*(value: string): Option[COLORREF] =
  ## Parse a #RRGGBB color string into a COLORREF.
  let trimmed = value.strip()
  if trimmed.len != 7 or trimmed[0] != '#':
    return

  try:
    let r = fromHex(trimmed[1 .. 2])
    let g = fromHex(trimmed[3 .. 4])
    let b = fromHex(trimmed[5 .. 6])
    some(COLORREF(RGB(r, g, b)))
  except ValueError:
    none(COLORREF)

proc colorRefToHex*(color: COLORREF): string =
  let r = int(color and 0xFF)
  let g = int((color shr 8) and 0xFF)
  let b = int((color shr 16) and 0xFF)
  "#" & r.toHex(2) & g.toHex(2) & b.toHex(2)

proc highlightElementBounds*(element: ptr IUIAutomationElement; color: COLORREF;
    durationMs: int = 1200; logger: Logger = nil): bool =
  ## Highlight a UIA element using its bounding rectangle.
  if element.isNil:
    if logger != nil:
      logger.warn("Highlight requested with nil UIA element")
    return false

  try:
    if element.isOffscreen():
      if logger != nil:
        logger.warn("Highlight skipped; element is offscreen")
      return false
  except CatchableError:
    discard

  let bounds = safeBoundingRect(element)
  if bounds.isNone:
    if logger != nil:
      logger.warn("Highlight skipped; bounding rectangle unavailable")
    return false

  let (left, top, width, height) = bounds.get()
  if width <= 1 or height <= 1:
    if logger != nil:
      logger.warn("Highlight skipped; zero-sized bounding box",
        [("width", $width), ("height", $height)])
    return false

  ensureOverlayClass(logger)
  if overlayClassAtom == 0:
    return false

  let hInstance = GetModuleHandleW(nil)
  let hwnd = CreateWindowExW(
    WS_EX_LAYERED or WS_EX_TRANSPARENT or WS_EX_TOPMOST or WS_EX_TOOLWINDOW or
      WS_EX_NOACTIVATE,
    overlayClassName,
    nil,
    WS_POPUP,
    left.int32,
    top.int32,
    width.int32 + 2,   # Slight padding so the border is visible
    height.int32 + 2,
    HWND(0),
    HMENU(0),
    hInstance,
    nil
  )

  if hwnd == 0:
    if logger != nil:
      logger.error("Failed to create highlight overlay window")
    return false

  discard SetWindowLongPtrW(hwnd, GWLP_USERDATA, cast[LONG_PTR](color))
  discard SetLayeredWindowAttributes(hwnd, overlayColorKey, 220.byte,
    LWA_COLORKEY or LWA_ALPHA)
  discard SetTimer(hwnd, overlayTimerId, UINT(durationMs), nil)

  discard ShowWindow(hwnd, SW_SHOWNOACTIVATE)
  discard UpdateWindow(hwnd)
  true
