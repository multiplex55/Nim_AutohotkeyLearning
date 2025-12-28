when system.hostOS != "windows":
  {.error: "UIA inspector window is only supported on Windows.".}

import std/[math, options, os, sets, strformat, strutils, tables]

import winim/lean
import winim/inc/commctrl
import winim/inc/uiautomation

import ../../core/logging
import ../uia/uia
import ./highlight_overlay
import ./state

const
  inspectorClassName = "NimUiaInspectorWindow"
  headerHeight = 48
  headerPadding = 8
  contentPadding = 8
  buttonHeight = 26
  buttonSpacing = 6
  sashWidth = 8
  minPanelWidth = 160
  expandTimerId = 99'u

  idInvoke = 1001
  idSetFocus = 1002
  idHighlight = 1003
  idCloseElement = 1004
  idExpandAll = 1005

  idMenuHighlightColor = 2001

type
  InspectorWindow = ref object
    hwnd: HWND
    header: HWND
    leftTree: HWND
    rightTree: HWND
    sashRect: RECT
    sashDragging: bool
    dragStartX: int
    dragStartWidth: int
    lastFocus: HWND
    uia: Uia
    logger: Logger
    nodes: Table[HTREEITEM, ptr IUIAutomationElement]
    statePath: string
    state: InspectorState
    highlightColor: COLORREF
    expandQueue: seq[HTREEITEM]
    expandActive: bool
    btnInvoke: HWND
    btnFocus: HWND
    btnHighlight: HWND
    btnClose: HWND
    btnExpand: HWND

var inspectors = initTable[HWND, InspectorWindow]()
var commonControlsReady = false
proc lParamX(lp: LPARAM): int =
  cast[int16](LOWORD(DWORD(lp))).int

proc lParamY(lp: LPARAM): int =
  cast[int16](HIWORD(DWORD(lp))).int

proc safeCurrentName(element: ptr IUIAutomationElement): string =
  try:
    element.currentName()
  except CatchableError:
    ""

proc safeAutomationId(element: ptr IUIAutomationElement): string =
  try:
    element.currentAutomationId()
  except CatchableError:
    ""

proc safeClassName(element: ptr IUIAutomationElement): string =
  try:
    element.currentClassName()
  except CatchableError:
    ""

proc safeControlType(element: ptr IUIAutomationElement): int =
  try:
    element.currentControlType()
  except CatchableError:
    -1

proc safeRuntimeId(element: ptr IUIAutomationElement): string =
  var arr: ptr SAFEARRAY = nil
  let hr = element.GetRuntimeId(addr arr)
  if FAILED(hr) or arr.isNil:
    return ""

  defer:
    if arr != nil:
      discard SafeArrayDestroy(arr)

  var lbound, ubound: LONG
  if FAILED(SafeArrayGetLBound(arr, 1, addr lbound)) or FAILED(SafeArrayGetUBound(arr,
      1, addr ubound)):
    return ""

  var parts: seq[string] = @[]
  var idx = lbound
  while idx <= ubound:
    var value: LONG
    if SUCCEEDED(SafeArrayGetElement(arr, addr idx, addr value)):
      parts.add($value)
    else:
      parts.add("?")
    inc idx

  "[" & parts.join(",") & "]"

proc controlTypeName(typeId: int): string =
  case typeId
  of UIA_ButtonControlTypeId: "Button"
  of UIA_CalendarControlTypeId: "Calendar"
  of UIA_CheckBoxControlTypeId: "CheckBox"
  of UIA_ComboBoxControlTypeId: "ComboBox"
  of UIA_DataGridControlTypeId: "DataGrid"
  of UIA_DocumentControlTypeId: "Document"
  of UIA_EditControlTypeId: "Edit"
  of UIA_GroupControlTypeId: "Group"
  of UIA_HyperlinkControlTypeId: "Hyperlink"
  of UIA_ImageControlTypeId: "Image"
  of UIA_ListControlTypeId: "List"
  of UIA_ListItemControlTypeId: "ListItem"
  of UIA_MenuControlTypeId: "Menu"
  of UIA_MenuBarControlTypeId: "MenuBar"
  of UIA_MenuItemControlTypeId: "MenuItem"
  of UIA_PaneControlTypeId: "Pane"
  of UIA_ProgressBarControlTypeId: "ProgressBar"
  of UIA_RadioButtonControlTypeId: "RadioButton"
  of UIA_ScrollBarControlTypeId: "ScrollBar"
  of UIA_SplitButtonControlTypeId: "SplitButton"
  of UIA_StatusBarControlTypeId: "StatusBar"
  of UIA_TabControlTypeId: "Tab"
  of UIA_TabItemControlTypeId: "TabItem"
  of UIA_TextControlTypeId: "Text"
  of UIA_TitleBarControlTypeId: "TitleBar"
  of UIA_ToolBarControlTypeId: "ToolBar"
  of UIA_ToolTipControlTypeId: "ToolTip"
  of UIA_TreeControlTypeId: "Tree"
  of UIA_TreeItemControlTypeId: "TreeItem"
  of UIA_WindowControlTypeId: "Window"
  else: fmt"ControlType({typeId})"

proc nodeLabel(element: ptr IUIAutomationElement): string =
  let name = safeCurrentName(element)
  let automationId = safeAutomationId(element)
  let ctrlType = controlTypeName(safeControlType(element))

  var parts: seq[string] = @[ctrlType]
  if name.len > 0:
    parts.add(&"\"{name}\"")
  if automationId.len > 0:
    parts.add(fmt"[{automationId}]")
  parts.join(" ")

proc releaseNodes(inspector: InspectorWindow) =
  var released = initHashSet[ptr IUIAutomationElement]()
  for _, element in inspector.nodes.pairs():
    if element == nil or element in released:
      continue
    released.incl(element)
    discard element.Release()
  inspector.nodes.clear()

proc addTreeItem(tree: HWND; parent: HTREEITEM; text: string;
    data: LPARAM = 0): HTREEITEM =
  var insert: TVINSERTSTRUCTW
  insert.hParent = parent
  insert.hInsertAfter = TVI_LAST
  insert.item.mask = UINT(TVIF_TEXT or TVIF_PARAM)
  let wide = newWideCString(text)
  insert.item.pszText = cast[LPWSTR](wide)
  insert.item.cchTextMax = int32(text.len)
  insert.item.lParam = data
  TreeView_InsertItemW(tree, addr insert)

proc addChildren(inspector: InspectorWindow; walker: ptr IUIAutomationTreeWalker;
    element: ptr IUIAutomationElement; parentItem: HTREEITEM; depth, maxDepth: int) =
  if depth >= maxDepth:
    return

  var child: ptr IUIAutomationElement
  let hrFirst = walker.GetFirstChildElement(element, addr child)
  if FAILED(hrFirst) or child.isNil:
    return

  var current = child
  while current != nil:
    discard current.AddRef()
    let item = addTreeItem(inspector.leftTree, parentItem, nodeLabel(current),
      cast[LPARAM](current))
    inspector.nodes[item] = current
    addChildren(inspector, walker, current, item, depth + 1, maxDepth)
    discard current.Release()

    var next: ptr IUIAutomationElement
    let hrNext = walker.GetNextSiblingElement(current, addr next)
    if FAILED(hrNext) or hrNext == S_FALSE:
      break
    current = next

proc rebuildElementTree(inspector: InspectorWindow) =
  TreeView_DeleteAllItems(inspector.leftTree)
  releaseNodes(inspector)

  if inspector.uia.isNil:
    if inspector.logger != nil:
      inspector.logger.error("UIA inspector missing automation instance")
    return

  var walker: ptr IUIAutomationTreeWalker
  let hrWalker = inspector.uia.automation.get_RawViewWalker(addr walker)
  if FAILED(hrWalker) or walker.isNil:
    if inspector.logger != nil:
      inspector.logger.error("Failed to create UIA walker",
        [("hresult", fmt"0x{hrWalker:X}")])
    return
  defer: discard walker.Release()

  var root = inspector.uia.rootElement()
  if root.isNil:
    if inspector.logger != nil:
      inspector.logger.warn("UIA root element unavailable; cannot build inspector tree")
    return

  discard root.AddRef()
  let rootItem = addTreeItem(inspector.leftTree, TVI_ROOT, nodeLabel(root),
    cast[LPARAM](root))
  inspector.nodes[rootItem] = root
  addChildren(inspector, walker, root, rootItem, 0, 4)
  TreeView_Expand(inspector.leftTree, rootItem, UINT(TVE_EXPAND))
  discard TreeView_SelectItem(inspector.leftTree, rootItem)
  populateProperties(inspector, root)

proc addPropertyEntry(tree: HWND; parent: HTREEITEM; key, value: string) =
  addTreeItem(tree, parent, fmt"{key}: {value}")

proc populateProperties(inspector: InspectorWindow; element: ptr IUIAutomationElement) =
  TreeView_DeleteAllItems(inspector.rightTree)
  if element.isNil:
    discard EnableWindow(inspector.btnInvoke, FALSE)
    discard EnableWindow(inspector.btnFocus, FALSE)
    discard EnableWindow(inspector.btnClose, FALSE)
    discard EnableWindow(inspector.btnHighlight, FALSE)
    discard EnableWindow(inspector.btnExpand, FALSE)
    return

  discard EnableWindow(inspector.btnInvoke, TRUE)
  discard EnableWindow(inspector.btnFocus, TRUE)
  discard EnableWindow(inspector.btnClose, TRUE)
  discard EnableWindow(inspector.btnHighlight, TRUE)

  let root = addTreeItem(inspector.rightTree, TVI_ROOT, "Current element")

  let identity = addTreeItem(inspector.rightTree, root, "Identity")
  addPropertyEntry(inspector.rightTree, identity, "Name", safeCurrentName(element))
  addPropertyEntry(inspector.rightTree, identity, "AutomationId", safeAutomationId(element))
  addPropertyEntry(inspector.rightTree, identity, "ClassName", safeClassName(element))
  addPropertyEntry(inspector.rightTree, identity, "ControlType", controlTypeName(safeControlType(element)))
  addPropertyEntry(inspector.rightTree, identity, "RuntimeId", safeRuntimeId(element))

  let stateGroup = addTreeItem(inspector.rightTree, root, "State")
  try:
    addPropertyEntry(inspector.rightTree, stateGroup, "IsEnabled", $element.isEnabled())
    addPropertyEntry(inspector.rightTree, stateGroup, "HasKeyboardFocus", $element.hasKeyboardFocus())
    addPropertyEntry(inspector.rightTree, stateGroup, "IsOffscreen", $element.isOffscreen())
    addPropertyEntry(inspector.rightTree, stateGroup, "IsControlElement", $element.isControlElement())
    addPropertyEntry(inspector.rightTree, stateGroup, "IsContentElement", $element.isContentElement())
  except CatchableError as exc:
    if inspector.logger != nil:
      inspector.logger.warn("Failed to read element state", [("error", exc.msg)])

  let boundsGroup = addTreeItem(inspector.rightTree, root, "Bounds")
  let bounds = safeBoundingRect(element)
  if bounds.isSome:
    let (left, top, width, height) = bounds.get()
    addPropertyEntry(inspector.rightTree, boundsGroup, "Left", $left.int)
    addPropertyEntry(inspector.rightTree, boundsGroup, "Top", $top.int)
    addPropertyEntry(inspector.rightTree, boundsGroup, "Width", $width.int)
    addPropertyEntry(inspector.rightTree, boundsGroup, "Height", $height.int)
  else:
    addPropertyEntry(inspector.rightTree, boundsGroup, "BoundingRectangle", "Unavailable")

  discard EnableWindow(inspector.btnExpand, TRUE)
  TreeView_Expand(inspector.rightTree, root, UINT(TVE_EXPAND))

proc currentSelection(inspector: InspectorWindow): ptr IUIAutomationElement =
  let selected = TreeView_GetSelection(inspector.leftTree)
  if selected == 0 or selected notin inspector.nodes:
    return nil
  inspector.nodes[selected]

proc handleInvoke(inspector: InspectorWindow) =
  let element = inspector.currentSelection()
  if element.isNil:
    return
  try:
    inspector.uia.invoke(element)
    if inspector.logger != nil:
      inspector.logger.info("Invoked UIA element",
        [("name", safeCurrentName(element)), ("automationId", safeAutomationId(element))])
  except CatchableError as exc:
    if inspector.logger != nil:
      inspector.logger.error("UIA invoke failed", [("error", exc.msg)])

proc handleSetFocus(inspector: InspectorWindow) =
  let element = inspector.currentSelection()
  if element.isNil:
    return
  try:
    let hr = element.SetFocus()
    if FAILED(hr):
      raise newException(UiaError, fmt"SetFocus failed (0x{hr:X})")
    if inspector.logger != nil:
      inspector.logger.info("Set keyboard focus to element",
        [("name", safeCurrentName(element)), ("automationId", safeAutomationId(element))])
  except CatchableError as exc:
    if inspector.logger != nil:
      inspector.logger.error("Failed to set focus", [("error", exc.msg)])

proc handleClose(inspector: InspectorWindow) =
  let element = inspector.currentSelection()
  if element.isNil:
    return
  try:
    inspector.uia.closeWindow(element)
    if inspector.logger != nil:
      inspector.logger.info("Issued close request to element",
        [("name", safeCurrentName(element)), ("automationId", safeAutomationId(element))])
  except CatchableError as exc:
    if inspector.logger != nil:
      inspector.logger.error("Failed to close element", [("error", exc.msg)])

proc handleHighlight(inspector: InspectorWindow) =
  let element = inspector.currentSelection()
  if element.isNil:
    discard EnableWindow(inspector.btnHighlight, FALSE)
    return

  if not highlightElementBounds(element, inspector.highlightColor, 1500, inspector.logger):
    if inspector.logger != nil:
      inspector.logger.warn("Highlight request failed for selected element")

proc collectExpandQueue(inspector: InspectorWindow) =
  inspector.expandQueue.setLen(0)
  var item = TreeView_GetRoot(inspector.rightTree)
  while item != 0:
    inspector.expandQueue.add(item)
    item = TreeView_GetNextSibling(inspector.rightTree, item)

proc beginExpandAll(inspector: InspectorWindow) =
  if inspector.expandActive:
    return
  collectExpandQueue(inspector)
  if inspector.expandQueue.len == 0:
    if inspector.logger != nil:
      inspector.logger.warn("Expand-all requested with no properties to expand")
    return
  inspector.expandActive = true
  discard EnableWindow(inspector.btnExpand, FALSE)
  discard SetTimer(inspector.hwnd, expandTimerId, 10, nil)

proc handleExpandTimer(inspector: InspectorWindow) =
  var processed = 0
  while inspector.expandQueue.len > 0 and processed < 64:
    let item = inspector.expandQueue.pop()
    discard TreeView_Expand(inspector.rightTree, item, UINT(TVE_EXPAND))
    var child = TreeView_GetChild(inspector.rightTree, item)
    while child != 0:
      inspector.expandQueue.add(child)
      child = TreeView_GetNextSibling(inspector.rightTree, child)
    inc processed

  if inspector.expandQueue.len == 0:
    KillTimer(inspector.hwnd, expandTimerId)
    inspector.expandActive = false
    discard EnableWindow(inspector.btnExpand, TRUE)

proc layoutHeader(inspector: InspectorWindow; width: int) =
  MoveWindow(inspector.header, 0, 0, width, headerHeight, TRUE)

  var x = contentPadding
  let y = headerPadding
  let btnWidths = [90, 96, 132, 120, 80]
  let buttons = [inspector.btnInvoke, inspector.btnFocus, inspector.btnHighlight,
      inspector.btnExpand, inspector.btnClose]
  for i, btn in buttons:
    MoveWindow(btn, x, y, btnWidths[i], buttonHeight, TRUE)
    x += btnWidths[i] + buttonSpacing

proc layoutContent(inspector: InspectorWindow; width, height: int) =
  var sashLeft = inspector.state.sashWidth
  let available = width - sashWidth
  if sashLeft <= 0 or sashLeft >= available - minPanelWidth:
    sashLeft = max(available div 2, minPanelWidth)
  sashLeft = clamp(sashLeft, minPanelWidth, available - minPanelWidth)
  inspector.state.sashWidth = sashLeft

  let contentTop = headerHeight
  let contentHeight = max(0, height - headerHeight)

  inspector.sashRect = RECT(left: LONG(sashLeft), top: LONG(contentTop),
    right: LONG(sashLeft + sashWidth), bottom: LONG(contentTop + contentHeight))

  let leftWidth = sashLeft - contentPadding
  let rightWidth = width - (sashLeft + sashWidth + contentPadding)
  MoveWindow(inspector.leftTree, contentPadding, contentTop + contentPadding,
    max(leftWidth - contentPadding, minPanelWidth - contentPadding),
    max(contentHeight - 2 * contentPadding, 100), TRUE)

  MoveWindow(inspector.rightTree, sashLeft + sashWidth + contentPadding,
    contentTop + contentPadding,
    max(rightWidth - contentPadding, minPanelWidth - contentPadding),
    max(contentHeight - 2 * contentPadding, 100), TRUE)
  discard InvalidateRect(inspector.hwnd, addr inspector.sashRect, FALSE)

proc applyLayout(inspector: InspectorWindow) =
  var rect: RECT
  discard GetClientRect(inspector.hwnd, addr rect)
  let width = rect.right - rect.left
  let height = rect.bottom - rect.top
  layoutHeader(inspector, width)
  layoutContent(inspector, width, height)

proc inspectorFromWindow(hwnd: HWND): InspectorWindow =
  if hwnd in inspectors:
    return inspectors[hwnd]
  let ptrVal = cast[InspectorWindow](GetWindowLongPtr(hwnd, GWLP_USERDATA))
  ptrVal

proc rebuildProperties(inspector: InspectorWindow) =
  let element = inspector.currentSelection()
  populateProperties(inspector, element)

proc createControls(inspector: InspectorWindow) =
  inspector.header = CreateWindowExW(
    0,
    WC_STATIC,
    nil,
    WS_CHILD or WS_VISIBLE,
    0, 0, 0, 0,
    inspector.hwnd,
    HMENU(0),
    GetModuleHandleW(nil),
    nil
  )

  let font = GetStockObject(DEFAULT_GUI_FONT)

  inspector.btnInvoke = CreateWindowExW(0, WC_BUTTON, newWideCString("Invoke"),
    WS_CHILD or WS_VISIBLE or WS_TABSTOP,
    0, 0, 0, 0, inspector.header, cast[HMENU](idInvoke),
    GetModuleHandleW(nil), nil)
  discard SendMessage(inspector.btnInvoke, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.btnFocus = CreateWindowExW(0, WC_BUTTON, newWideCString("Set Focus"),
    WS_CHILD or WS_VISIBLE or WS_TABSTOP,
    0, 0, 0, 0, inspector.header, cast[HMENU](idSetFocus),
    GetModuleHandleW(nil), nil)
  discard SendMessage(inspector.btnFocus, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.btnHighlight = CreateWindowExW(0, WC_BUTTON,
    newWideCString("Highlight selected"),
    WS_CHILD or WS_VISIBLE or WS_TABSTOP,
    0, 0, 0, 0, inspector.header, cast[HMENU](idHighlight),
    GetModuleHandleW(nil), nil)
  discard SendMessage(inspector.btnHighlight, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.btnExpand = CreateWindowExW(0, WC_BUTTON,
    newWideCString("Expand All Current"),
    WS_CHILD or WS_VISIBLE or WS_TABSTOP,
    0, 0, 0, 0, inspector.header, cast[HMENU](idExpandAll),
    GetModuleHandleW(nil), nil)
  discard SendMessage(inspector.btnExpand, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.btnClose = CreateWindowExW(0, WC_BUTTON, newWideCString("Close"),
    WS_CHILD or WS_VISIBLE or WS_TABSTOP,
    0, 0, 0, 0, inspector.header, cast[HMENU](idCloseElement),
    GetModuleHandleW(nil), nil)
  discard SendMessage(inspector.btnClose, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  var treeStyle = WS_CHILD or WS_VISIBLE or WS_TABSTOP or TVS_HASBUTTONS or
      TVS_LINESATROOT or TVS_HASLINES or WS_BORDER
  inspector.leftTree = CreateWindowExW(WS_EX_CLIENTEDGE, WC_TREEVIEWW, nil,
    treeStyle, 0, 0, 0, 0, inspector.hwnd, HMENU(0), GetModuleHandleW(nil), nil)
  discard SendMessage(inspector.leftTree, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.rightTree = CreateWindowExW(WS_EX_CLIENTEDGE, WC_TREEVIEWW, nil,
    treeStyle, 0, 0, 0, 0, inspector.hwnd, HMENU(0), GetModuleHandleW(nil), nil)
  discard SendMessage(inspector.rightTree, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  discard EnableWindow(inspector.btnInvoke, FALSE)
  discard EnableWindow(inspector.btnFocus, FALSE)
  discard EnableWindow(inspector.btnClose, FALSE)
  discard EnableWindow(inspector.btnHighlight, FALSE)
  discard EnableWindow(inspector.btnExpand, FALSE)

  rebuildElementTree(inspector)
  applyLayout(inspector)

proc updateHighlightColor(inspector: InspectorWindow; hexColor: string) =
  let parsed = parseColorRef(hexColor)
  if parsed.isSome:
    inspector.highlightColor = parsed.get()
  else:
    inspector.highlightColor = COLORREF(RGB(255, 0, 0))
  inspector.state.highlightColor = colorRefToHex(inspector.highlightColor)

proc registerMenu(inspector: InspectorWindow) =
  let menuBar = CreateMenu()
  let viewMenu = CreateMenu()
  discard AppendMenuW(viewMenu, MF_STRING, idMenuHighlightColor,
    newWideCString("Highlight Color..."))
  discard AppendMenuW(menuBar, MF_POPUP, cast[UINT_PTR](viewMenu),
    newWideCString("&View"))
  discard SetMenu(inspector.hwnd, menuBar)

proc handleCommand(inspector: InspectorWindow; wParam: WPARAM) =
  let id = int(LOWORD(DWORD(wParam)))
  case id
  of idInvoke:
    handleInvoke(inspector)
  of idSetFocus:
    handleSetFocus(inspector)
  of idHighlight:
    handleHighlight(inspector)
  of idCloseElement:
    handleClose(inspector)
  of idExpandAll:
    beginExpandAll(inspector)
  of idMenuHighlightColor:
    var cc: CHOOSECOLORW
    cc.lStructSize = DWORD(sizeof(cc))
    cc.hwndOwner = inspector.hwnd
    cc.Flags = CC_FULLOPEN or CC_RGBINIT
    cc.rgbResult = inspector.highlightColor
    var custom: array[16, COLORREF]
    cc.lpCustColors = cast[LPCOLORREF](addr custom[0])
    if ChooseColorW(addr cc) != 0:
      inspector.highlightColor = cc.rgbResult
      inspector.state.highlightColor = colorRefToHex(cc.rgbResult)
      saveInspectorState(inspector.statePath, inspector.state, inspector.logger)
  else:
    discard

proc handleNotify(inspector: InspectorWindow; lParam: LPARAM) =
  let hdr = cast[ptr NMHDR](lParam)
  if hdr.hwndFrom == inspector.leftTree and hdr.code == uint(TVN_SELCHANGEDW):
    let info = cast[ptr NMTREEVIEWW](lParam)
    populateProperties(inspector, inspector.nodes.getOrDefault(info.itemNew.hItem, nil))

var inspectorClassRegistered = false

proc registerInspectorClass() =
  if inspectorClassRegistered:
    return
  var wc: WNDCLASSEXW
  wc.cbSize = UINT(sizeof(wc))
  wc.style = CS_HREDRAW or CS_VREDRAW
  wc.lpfnWndProc = cast[WNDPROC](proc(hwnd: HWND; msg: UINT; wParam: WPARAM;
      lParam: LPARAM): LRESULT {.stdcall.} =
    var inspector = inspectorFromWindow(hwnd)

    case msg
    of WM_NCCREATE:
      let cs = cast[ptr CREATESTRUCT](lParam)
      discard SetWindowLongPtr(hwnd, GWLP_USERDATA, cast[LONG_PTR](cs.lpCreateParams))
      return DefWindowProc(hwnd, msg, wParam, lParam)
    of WM_CREATE:
      inspector = inspectorFromWindow(hwnd)
      if inspector != nil:
        inspector.hwnd = hwnd
        registerMenu(inspector)
        createControls(inspector)
      return 0
    of WM_SIZE:
      if inspector != nil:
        applyLayout(inspector)
      return 0
    of WM_LBUTTONDOWN:
      if inspector != nil:
        let x = lParamX(lParam)
        let y = lParamY(lParam)
        if x >= inspector.sashRect.left and x <= inspector.sashRect.right and
            y >= inspector.sashRect.top and y <= inspector.sashRect.bottom:
          inspector.sashDragging = true
          inspector.dragStartX = int(x)
          inspector.dragStartWidth = inspector.state.sashWidth
          inspector.lastFocus = GetFocus()
          discard SetCapture(hwnd)
      return 0
    of WM_MOUSEMOVE:
      if inspector != nil and inspector.sashDragging:
        let x = lParamX(lParam)
        let delta = int(x) - inspector.dragStartX
        inspector.state.sashWidth = inspector.dragStartWidth + delta
        applyLayout(inspector)
      return 0
    of WM_LBUTTONUP:
      if inspector != nil and inspector.sashDragging:
        inspector.sashDragging = false
        discard ReleaseCapture()
        applyLayout(inspector)
        saveInspectorState(inspector.statePath, inspector.state, inspector.logger)
        if inspector.lastFocus != HWND(0):
          discard SetFocus(inspector.lastFocus)
      return 0
    of WM_COMMAND:
      if inspector != nil:
        handleCommand(inspector, wParam)
      return 0
    of WM_NOTIFY:
      if inspector != nil:
        handleNotify(inspector, lParam)
      return 0
    of WM_PAINT:
      if inspector != nil:
        var ps: PAINTSTRUCT
        let hdc = BeginPaint(hwnd, addr ps)
        var clientRect: RECT
        discard GetClientRect(hwnd, addr clientRect)
        let bg = GetSysColorBrush(COLOR_WINDOW)
        discard FillRect(hdc, addr clientRect, bg)
        let sashBrush = CreateSolidBrush(RGB(230, 230, 230))
        discard FillRect(hdc, addr inspector.sashRect, sashBrush)
        discard DeleteObject(sashBrush)
        discard EndPaint(hwnd, addr ps)
        return 0
    of WM_TIMER:
      if inspector != nil and UINT(wParam) == expandTimerId:
        handleExpandTimer(inspector)
      return 0
    of WM_DESTROY:
      if inspector != nil:
        releaseNodes(inspector)
        saveInspectorState(inspector.statePath, inspector.state, inspector.logger)
        KillTimer(hwnd, expandTimerId)
        if inspector.hwnd in inspectors:
          inspectors.del(inspector.hwnd)
      return 0
    else:
      discard
    result = DefWindowProc(hwnd, msg, wParam, lParam)
  )
  wc.cbClsExtra = 0
  wc.cbWndExtra = 0
  wc.hInstance = GetModuleHandleW(nil)
  wc.hIcon = LoadIcon(0, IDI_APPLICATION)
  wc.hCursor = LoadCursor(0, IDC_ARROW)
  wc.hbrBackground = cast[HBRUSH](COLOR_WINDOW + 1)
  wc.lpszMenuName = nil
  wc.lpszClassName = inspectorClassName
  if RegisterClassExW(addr wc) != 0:
    inspectorClassRegistered = true

proc ensureCommonControls() =
  if commonControlsReady:
    return
  var icc: INITCOMMONCONTROLSEX
  icc.dwSize = DWORD(sizeof(icc))
  icc.dwICC = ICC_TREEVIEW_CLASSES
  discard InitCommonControlsEx(addr icc)
  commonControlsReady = true

proc defaultInspectorStatePath(): string =
  joinPath(getCurrentDir(), DEFAULT_INSPECTOR_STATE_FILENAME)

proc showInspectorWindow*(uia: Uia; logger: Logger = nil;
    statePath: string = ""): bool =
  ensureCommonControls()
  registerInspectorClass()

  var insp = InspectorWindow(
    uia: uia,
    logger: logger,
    statePath: if statePath.len > 0: statePath else: defaultInspectorStatePath()
  )
  insp.state = loadInspectorState(insp.statePath, logger)
  updateHighlightColor(insp, insp.state.highlightColor)

  let hwnd = CreateWindowExW(
    WS_EX_APPWINDOW,
    inspectorClassName,
    newWideCString("UIA Inspector"),
    WS_OVERLAPPEDWINDOW or WS_VISIBLE,
    CW_USEDEFAULT, CW_USEDEFAULT,
    1180, 760,
    HWND(0),
    HMENU(0),
    GetModuleHandleW(nil),
    cast[LPVOID](insp)
  )

  if hwnd == 0:
    if logger != nil:
      logger.error("Failed to create inspector window")
    return false

  insp.hwnd = hwnd
  inspectors[hwnd] = insp
  discard ShowWindow(hwnd, SW_SHOWNORMAL)
  discard UpdateWindow(hwnd)
  true

proc focusExistingInspector*(): bool =
  for hwnd in inspectors.keys():
    if IsWindow(hwnd) != 0:
      discard ShowWindow(hwnd, SW_SHOWNORMAL)
      discard SetForegroundWindow(hwnd)
      return true
  false
