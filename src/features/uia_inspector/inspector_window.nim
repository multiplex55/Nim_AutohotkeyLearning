when system.hostOS != "windows":
  {.error: "UIA inspector window is only supported on Windows.".}

import std/[options, os, sets, strformat, strutils, tables]

import winim/lean
import winim/com
import winim/inc/commctrl
import winim/inc/commdlg
import winim/inc/uiautomation
import winim/inc/winver

import ../../core/logging
import ../uia/uia
import ./highlight_overlay
import ./state

const
  inspectorClassName = "NimUiaInspectorWindow"
  contentPadding = 8
  groupPadding = 8
  buttonHeight = 26
  buttonSpacing = 6
  splitterWidth = 8
  minPanelWidth = 180
  minMiddleHeight = 100
  statusBarHeight = 24
  bottomPadding = 6
  expandTimerId = UINT_PTR(99)

  idInvoke = 1001
  idSetFocus = 1002
  idHighlight = 1003
  idCloseElement = 1004
  idExpandAll = 1005
  idRefresh = 1006

  idMenuHighlightColor = 2001

type
  InspectorWindow = ref object
    hwnd: HWND
    gbWindowList: HWND
    gbWindowInfo: HWND
    gbProperties: HWND
    gbPatterns: HWND
    windowFilterLabel: HWND
    windowFilterEdit: HWND
    windowClassLabel: HWND
    windowClassEdit: HWND
    btnRefresh: HWND
    windowList: HWND
    windowInfoText: HWND
    propertiesTree: HWND
    patternsTree: HWND
    mainTree: HWND
    statusBar: HWND
    splitters: array[3, RECT]
    splitterDragging: int
    dragStartX: int
    dragStartY: int
    dragStartLeft: int
    dragStartMiddle: int
    dragStartProperties: int
    lastFocus: HWND
    uia: Uia
    logger: Logger
    nodes: Table[HTREEITEM, ptr IUIAutomationElement]
    statePath: string
    state: InspectorState
    highlightColor: COLORREF
    uiaVersion: string
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

proc fileVersion(path: string): string =
  var handle: DWORD
  let size = GetFileVersionInfoSizeW(newWideCString(path), addr handle)
  if size == 0:
    return ""

  var buffer = newSeq[BYTE](int(size))
  if GetFileVersionInfoW(newWideCString(path), 0, size, addr buffer[0]) == 0:
    return ""

  var info: ptr VS_FIXEDFILEINFO
  var len: UINT
  if VerQueryValueW(addr buffer[0], newWideCString("\\"), cast[ptr LPVOID](addr info),
      addr len) == 0:
    return ""

  if info.isNil:
    return ""
  let major = HIWORD(info.dwFileVersionMS)
  let minor = LOWORD(info.dwFileVersionMS)
  let build = HIWORD(info.dwFileVersionLS)
  let revision = LOWORD(info.dwFileVersionLS)
  fmt"{major}.{minor}.{build}.{revision}"

proc getUiaCoreVersion(): string =
  let moduleName = newWideCString("uiautomationcore.dll")
  var module = GetModuleHandleW(moduleName)
  if module == 0:
    module = LoadLibraryW(moduleName)
  if module == 0:
    return "Unknown"

  var pathBuf: array[MAX_PATH, WCHAR]
  let written = GetModuleFileNameW(module, addr pathBuf[0], DWORD(pathBuf.len))
  if written == 0:
    return "Unknown"
  let path = fromWideCString(addr pathBuf[0])
  let ver = fileVersion(path)
  if ver.len > 0:
    ver
  else:
    "Unknown"

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

proc populateProperties(inspector: InspectorWindow; element: ptr IUIAutomationElement)

proc addTreeItem(tree: HWND; parent: HTREEITEM; text: string;
    data: LPARAM = 0): HTREEITEM =
  var insert: TVINSERTSTRUCTW
  insert.hParent = parent
  insert.hInsertAfter = TVI_LAST
  insert.item.mask = UINT(TVIF_TEXT or TVIF_PARAM)
  let wide = newWideCString(text)
  insert.item.pszText = wide
  insert.item.cchTextMax = int32(text.len)
  insert.item.lParam = data
  TreeView_InsertItem(tree, addr insert)

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
    let item = addTreeItem(inspector.mainTree, parentItem, nodeLabel(current),
      cast[LPARAM](current))
    inspector.nodes[item] = current
    addChildren(inspector, walker, current, item, depth + 1, maxDepth)

    var next: ptr IUIAutomationElement
    let hrNext = walker.GetNextSiblingElement(current, addr next)
    if FAILED(hrNext) or hrNext == S_FALSE:
      break
    current = next

proc rebuildElementTree(inspector: InspectorWindow) =
  TreeView_DeleteAllItems(inspector.mainTree)
  discard SendMessage(inspector.windowList, LVM_DELETEALLITEMS, 0, 0)
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

  let rootItem = addTreeItem(inspector.mainTree, TVI_ROOT, nodeLabel(root),
    cast[LPARAM](root))
  inspector.nodes[rootItem] = root
  var lvItem: LVITEMW
  lvItem.mask = LVIF_TEXT
  lvItem.pszText = newWideCString(nodeLabel(root))
  discard SendMessage(inspector.windowList, LVM_INSERTITEMW, 0, cast[LPARAM](addr lvItem))
  addChildren(inspector, walker, root, rootItem, 0, 4)
  TreeView_Expand(inspector.mainTree, rootItem, UINT(TVE_EXPAND))
  discard TreeView_SelectItem(inspector.mainTree, rootItem)
  populateProperties(inspector, root)

proc addPropertyEntry(tree: HWND; parent: HTREEITEM; key, value: string) =
  discard addTreeItem(tree, parent, fmt"{key}: {value}")

proc populateProperties(inspector: InspectorWindow; element: ptr IUIAutomationElement) =
  TreeView_DeleteAllItems(inspector.propertiesTree)
  TreeView_DeleteAllItems(inspector.patternsTree)
  if element.isNil:
    discard EnableWindow(inspector.btnInvoke, FALSE)
    discard EnableWindow(inspector.btnFocus, FALSE)
    discard EnableWindow(inspector.btnClose, FALSE)
    discard EnableWindow(inspector.btnHighlight, FALSE)
    discard EnableWindow(inspector.btnExpand, FALSE)
    discard SetWindowTextW(inspector.windowInfoText, newWideCString("No element selected"))
    discard addTreeItem(inspector.patternsTree, TVI_ROOT, "No patterns available")
    return

  discard EnableWindow(inspector.btnInvoke, TRUE)
  discard EnableWindow(inspector.btnFocus, TRUE)
  discard EnableWindow(inspector.btnClose, TRUE)
  discard EnableWindow(inspector.btnHighlight, TRUE)

  let infoText = fmt"Name: {safeCurrentName(element)}\nClass: {safeClassName(element)}\nAutomationId: {safeAutomationId(element)}"
  discard SetWindowTextW(inspector.windowInfoText, newWideCString(infoText))

  let root = addTreeItem(inspector.propertiesTree, TVI_ROOT, "Current element")

  let identity = addTreeItem(inspector.propertiesTree, root, "Identity")
  addPropertyEntry(inspector.propertiesTree, identity, "Name", safeCurrentName(element))
  addPropertyEntry(inspector.propertiesTree, identity, "AutomationId", safeAutomationId(element))
  addPropertyEntry(inspector.propertiesTree, identity, "ClassName", safeClassName(element))
  addPropertyEntry(inspector.propertiesTree, identity, "ControlType", controlTypeName(safeControlType(element)))

  let stateGroup = addTreeItem(inspector.propertiesTree, root, "State")
  try:
    addPropertyEntry(inspector.propertiesTree, stateGroup, "IsEnabled", $element.isEnabled())
    addPropertyEntry(inspector.propertiesTree, stateGroup, "HasKeyboardFocus", $element.hasKeyboardFocus())
    addPropertyEntry(inspector.propertiesTree, stateGroup, "IsOffscreen", $element.isOffscreen())
    addPropertyEntry(inspector.propertiesTree, stateGroup, "IsControlElement", $element.isControlElement())
    addPropertyEntry(inspector.propertiesTree, stateGroup, "IsContentElement", $element.isContentElement())
  except CatchableError as exc:
    if inspector.logger != nil:
      inspector.logger.warn("Failed to read element state", [("error", exc.msg)])

  let boundsGroup = addTreeItem(inspector.propertiesTree, root, "Bounds")
  let bounds = safeBoundingRect(element)
  if bounds.isSome:
    let (left, top, width, height) = bounds.get()
    addPropertyEntry(inspector.propertiesTree, boundsGroup, "Left", $left.int)
    addPropertyEntry(inspector.propertiesTree, boundsGroup, "Top", $top.int)
    addPropertyEntry(inspector.propertiesTree, boundsGroup, "Width", $width.int)
    addPropertyEntry(inspector.propertiesTree, boundsGroup, "Height", $height.int)
  else:
    addPropertyEntry(inspector.propertiesTree, boundsGroup, "BoundingRectangle", "Unavailable")

  discard EnableWindow(inspector.btnExpand, TRUE)
  TreeView_Expand(inspector.propertiesTree, root, UINT(TVE_EXPAND))
  discard addTreeItem(inspector.patternsTree, TVI_ROOT, "Patterns not yet listed")

proc currentSelection(inspector: InspectorWindow): ptr IUIAutomationElement =
  let selected = TreeView_GetSelection(inspector.mainTree)
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
  var item = TreeView_GetRoot(inspector.propertiesTree)
  while item != 0:
    inspector.expandQueue.add(item)
    item = TreeView_GetNextSibling(inspector.propertiesTree, item)

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
  discard SetTimer(inspector.hwnd, UINT_PTR(expandTimerId), UINT(10), nil)

proc handleExpandTimer(inspector: InspectorWindow) =
  var processed = 0
  while inspector.expandQueue.len > 0 and processed < 64:
    let item = inspector.expandQueue.pop()
    discard TreeView_Expand(inspector.propertiesTree, item, UINT(TVE_EXPAND))
    var child = TreeView_GetChild(inspector.propertiesTree, item)
    while child != 0:
      inspector.expandQueue.add(child)
      child = TreeView_GetNextSibling(inspector.propertiesTree, child)
    inc processed

  if inspector.expandQueue.len == 0:
    KillTimer(inspector.hwnd, UINT_PTR(expandTimerId))
    inspector.expandActive = false
    discard EnableWindow(inspector.btnExpand, TRUE)

proc layoutContent(inspector: InspectorWindow; width, height: int) =
  var sbHeight = statusBarHeight
  if inspector.statusBar != 0:
    var sbRect: RECT
    discard GetWindowRect(inspector.statusBar, addr sbRect)
    sbHeight = max(0, sbRect.bottom - sbRect.top)
  let usableHeight = max(0, height - sbHeight - bottomPadding)
  let minSection = max(minMiddleHeight div 2, 50)
  var leftWidth = inspector.state.leftWidth
  var middleWidth = inspector.state.middleWidth
  let minRightWidth = minPanelWidth
  let availableWidth = width - 2 * splitterWidth - 2 * contentPadding

  if leftWidth <= 0:
    leftWidth = minPanelWidth
  if middleWidth <= 0:
    middleWidth = minPanelWidth

  leftWidth = clamp(leftWidth, minPanelWidth, availableWidth - minPanelWidth - splitterWidth)
  middleWidth = clamp(middleWidth, minPanelWidth,
    availableWidth - leftWidth - splitterWidth - minRightWidth)

  var rightWidth = availableWidth - leftWidth - middleWidth - splitterWidth
  if rightWidth < minRightWidth:
    rightWidth = minRightWidth
    middleWidth = availableWidth - leftWidth - splitterWidth - rightWidth
    middleWidth = max(middleWidth, minPanelWidth)

  inspector.state.leftWidth = leftWidth
  inspector.state.middleWidth = middleWidth

  let leftX = contentPadding
  let middleX = leftX + leftWidth + splitterWidth
  let rightX = middleX + middleWidth + splitterWidth
  let contentTop = contentPadding

  inspector.splitters[0] = RECT(
    left: LONG(leftX + leftWidth),
    right: LONG(leftX + leftWidth + splitterWidth),
    top: LONG(contentTop),
    bottom: LONG(contentTop + usableHeight)
  )
  inspector.splitters[1] = RECT(
    left: LONG(middleX + middleWidth),
    right: LONG(middleX + middleWidth + splitterWidth),
    top: LONG(contentTop),
    bottom: LONG(contentTop + usableHeight)
  )

  MoveWindow(inspector.gbWindowList, leftX.cint, contentTop.cint,
    leftWidth.int32, usableHeight.int32, TRUE)

  let groupInnerLeft = leftX + groupPadding
  var currentY = contentTop + groupPadding
  MoveWindow(inspector.windowFilterLabel, groupInnerLeft.cint, currentY.cint, (leftWidth - 2 * groupPadding).int32, 16, TRUE)
  currentY += 18
  MoveWindow(inspector.windowFilterEdit, groupInnerLeft.cint, currentY.cint,
    (leftWidth - 2 * groupPadding).int32, 22, TRUE)
  currentY += 26
  MoveWindow(inspector.windowClassLabel, groupInnerLeft.cint, currentY.cint,
    (leftWidth - 2 * groupPadding).int32, 16, TRUE)
  currentY += 18
  MoveWindow(inspector.windowClassEdit, groupInnerLeft.cint, currentY.cint,
    (leftWidth - 2 * groupPadding).int32, 22, TRUE)
  currentY += 28
  MoveWindow(inspector.btnRefresh, groupInnerLeft.cint, currentY.cint, 96, buttonHeight.int32, TRUE)
  currentY += buttonHeight + groupPadding
  MoveWindow(inspector.windowList, groupInnerLeft.cint, currentY.cint,
    (leftWidth - 2 * groupPadding).int32,
    max(usableHeight - (currentY - contentTop) - groupPadding, 80).int32, TRUE)

  let middleHeight = usableHeight
  var infoHeight = inspector.state.infoHeight
  let maxInfo = max(minSection, middleHeight - splitterWidth - minSection)
  let minInfo = minSection
  infoHeight = clamp(infoHeight, minInfo, maxInfo)
  inspector.state.infoHeight = infoHeight
  MoveWindow(inspector.gbWindowInfo, middleX.cint, contentTop.cint,
    middleWidth.int32, infoHeight.int32, TRUE)

  let infoInnerWidth = middleWidth - 2 * groupPadding
  let infoInnerHeight = max(infoHeight - 2 * groupPadding, 40)
  MoveWindow(inspector.windowInfoText, (middleX + groupPadding).cint,
    (contentTop + groupPadding).cint, infoInnerWidth.int32,
    infoInnerHeight.int32, TRUE)

  let lowerAvailable = middleHeight - infoHeight - splitterWidth
  var propertiesHeight = inspector.state.propertiesHeight
  var patternsHeight = 0
  let availableSections = max(0, lowerAvailable - splitterWidth)
  if availableSections < minSection * 2:
    propertiesHeight = availableSections div 2
  else:
    propertiesHeight = clamp(propertiesHeight, minSection, availableSections - minSection)
  patternsHeight = max(availableSections - propertiesHeight, 0)
  inspector.state.propertiesHeight = propertiesHeight

  let propertiesY = contentTop + infoHeight + splitterWidth
  let propBoxHeight = max(propertiesHeight, 1)
  MoveWindow(inspector.gbProperties, middleX.cint, propertiesY.cint, middleWidth.int32,
    propBoxHeight.int32, TRUE)

  let propInnerY = propertiesY + groupPadding
  let propInnerHeight = propBoxHeight - 2 * groupPadding - buttonHeight - buttonSpacing
  MoveWindow(inspector.propertiesTree, (middleX + groupPadding).cint,
    propInnerY.cint, (middleWidth - 2 * groupPadding).int32,
    max(propInnerHeight, 20).int32, TRUE)

  let propButtonsY = max(propertiesY + propBoxHeight - groupPadding - buttonHeight,
    propertiesY + groupPadding)
  var buttonX = middleX + groupPadding
  MoveWindow(inspector.btnHighlight, buttonX.cint, propButtonsY.cint, 120, buttonHeight.int32, TRUE)
  buttonX += 120 + buttonSpacing
  MoveWindow(inspector.btnExpand, buttonX.cint, propButtonsY.cint, 140, buttonHeight.int32, TRUE)

  let splitterY = propertiesY + propertiesHeight
  inspector.splitters[2] = RECT(
    left: LONG(middleX),
    right: LONG(middleX + middleWidth),
    top: LONG(splitterY),
    bottom: LONG(splitterY + splitterWidth)
  )

  let patternsY = splitterY + splitterWidth
  let patternBoxHeight = max(patternsHeight, 1)
  MoveWindow(inspector.gbPatterns, middleX.cint, patternsY.cint,
    middleWidth.int32, patternBoxHeight.int32, TRUE)

  let patternsInnerHeight = max(patternBoxHeight - 2 * groupPadding - buttonHeight - buttonSpacing,
    20)
  MoveWindow(inspector.patternsTree, (middleX + groupPadding).cint,
    (patternsY + groupPadding).cint, (middleWidth - 2 * groupPadding).int32,
    patternsInnerHeight.int32, TRUE)

  let patternBtnY = max(patternsY + patternBoxHeight - groupPadding - buttonHeight,
    patternsY + groupPadding)
  var patternBtnX = middleX + groupPadding
  MoveWindow(inspector.btnInvoke, patternBtnX.cint, patternBtnY.cint, 100, buttonHeight.int32, TRUE)
  patternBtnX += 100 + buttonSpacing
  MoveWindow(inspector.btnFocus, patternBtnX.cint, patternBtnY.cint, 100, buttonHeight.int32, TRUE)
  patternBtnX += 100 + buttonSpacing
  MoveWindow(inspector.btnClose, patternBtnX.cint, patternBtnY.cint, 100, buttonHeight.int32, TRUE)

  MoveWindow(inspector.mainTree, rightX.cint, contentTop.cint, rightWidth.int32,
    usableHeight.int32, TRUE)

  MoveWindow(inspector.statusBar, 0, usableHeight.cint,
    width.int32, sbHeight.int32, TRUE)

  var parts: array[2, int32]
  parts[0] = max(0, width - 300).int32
  parts[1] = -1
  discard SendMessage(inspector.statusBar, SB_SETPARTS, WPARAM(parts.len),
    cast[LPARAM](addr parts[0]))
  updateStatusBar(inspector)
  discard InvalidateRect(inspector.hwnd, nil, FALSE)

proc applyLayout(inspector: InspectorWindow) =
  var rect: RECT
  discard GetClientRect(inspector.hwnd, addr rect)
  let width = rect.right - rect.left
  let height = rect.bottom - rect.top
  layoutContent(inspector, width, height)

proc inspectorFromWindow(hwnd: HWND): InspectorWindow =
  if hwnd in inspectors:
    return inspectors[hwnd]
  let ptrVal = cast[InspectorWindow](GetWindowLongPtr(hwnd, GWLP_USERDATA))
  ptrVal

proc createControls(inspector: InspectorWindow) =
  let font = GetStockObject(DEFAULT_GUI_FONT)
  let hInst = GetModuleHandleW(nil)

  inspector.gbWindowList = CreateWindowExW(0, WC_BUTTON,
    newWideCString("Window List"),
    WS_CHILD or WS_VISIBLE or BS_GROUPBOX,
    0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.gbWindowList, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowFilterLabel = CreateWindowExW(0, WC_STATIC,
    newWideCString("Title filter:"), WS_CHILD or WS_VISIBLE,
    0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.windowFilterLabel, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowFilterEdit = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or WS_TABSTOP or ES_AUTOHSCROLL,
    0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.windowFilterEdit, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowClassLabel = CreateWindowExW(0, WC_STATIC,
    newWideCString("Class filter:"), WS_CHILD or WS_VISIBLE,
    0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.windowClassLabel, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowClassEdit = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or WS_TABSTOP or ES_AUTOHSCROLL,
    0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.windowClassEdit, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.btnRefresh = CreateWindowExW(0, WC_BUTTON, newWideCString("Refresh"),
    WS_CHILD or WS_VISIBLE or WS_TABSTOP,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idRefresh), hInst, nil)
  discard SendMessage(inspector.btnRefresh, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  var listStyle = WS_CHILD or WS_VISIBLE or WS_TABSTOP or LVS_REPORT or LVS_SHOWSELALWAYS or
      LVS_SINGLESEL or WS_BORDER
  inspector.windowList = CreateWindowExW(DWORD(WS_EX_CLIENTEDGE), WC_LISTVIEWW, nil,
    DWORD(listStyle), 0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.windowList, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  discard SendMessage(inspector.windowList, LVM_SETEXTENDEDLISTVIEWSTYLE, 0,
    LPARAM(LVS_EX_FULLROWSELECT or LVS_EX_GRIDLINES))
  var col: LVCOLUMNW
  col.mask = LVCF_TEXT or LVCF_WIDTH
  col.cx = 200
  col.pszText = newWideCString("Window")
  discard SendMessage(inspector.windowList, LVM_INSERTCOLUMNW, 0, cast[LPARAM](addr col))

  inspector.gbWindowInfo = CreateWindowExW(0, WC_BUTTON, newWideCString("Window Info"),
    WS_CHILD or WS_VISIBLE or BS_GROUPBOX, 0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.gbWindowInfo, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowInfoText = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or ES_MULTILINE or ES_AUTOVSCROLL or ES_READONLY or WS_VSCROLL,
    0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.windowInfoText, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  discard SetWindowTextW(inspector.windowInfoText, newWideCString("No element selected"))

  inspector.gbProperties = CreateWindowExW(0, WC_BUTTON, newWideCString("Properties"),
    WS_CHILD or WS_VISIBLE or BS_GROUPBOX, 0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.gbProperties, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  var treeStyle = WS_CHILD or WS_VISIBLE or WS_TABSTOP or TVS_HASBUTTONS or
      TVS_LINESATROOT or TVS_HASLINES or WS_BORDER
  inspector.propertiesTree = CreateWindowExW(DWORD(WS_EX_CLIENTEDGE), WC_TREEVIEWW, nil,
    DWORD(treeStyle), 0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.propertiesTree, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.gbPatterns = CreateWindowExW(0, WC_BUTTON, newWideCString("Patterns"),
    WS_CHILD or WS_VISIBLE or BS_GROUPBOX, 0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.gbPatterns, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.patternsTree = CreateWindowExW(DWORD(WS_EX_CLIENTEDGE), WC_TREEVIEWW, nil,
    DWORD(treeStyle), 0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.patternsTree, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.btnHighlight = CreateWindowExW(0, WC_BUTTON,
    newWideCString("Highlight selected"),
    WS_CHILD or WS_VISIBLE or WS_TABSTOP,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idHighlight),
    hInst, nil)
  discard SendMessage(inspector.btnHighlight, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.btnExpand = CreateWindowExW(0, WC_BUTTON,
    newWideCString("Expand properties"),
    WS_CHILD or WS_VISIBLE or WS_TABSTOP,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idExpandAll),
    hInst, nil)
  discard SendMessage(inspector.btnExpand, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.btnInvoke = CreateWindowExW(0, WC_BUTTON, newWideCString("Invoke"),
    WS_CHILD or WS_VISIBLE or WS_TABSTOP,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idInvoke),
    hInst, nil)
  discard SendMessage(inspector.btnInvoke, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.btnFocus = CreateWindowExW(0, WC_BUTTON, newWideCString("Set Focus"),
    WS_CHILD or WS_VISIBLE or WS_TABSTOP,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idSetFocus),
    hInst, nil)
  discard SendMessage(inspector.btnFocus, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.btnClose = CreateWindowExW(0, WC_BUTTON, newWideCString("Close"),
    WS_CHILD or WS_VISIBLE or WS_TABSTOP,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idCloseElement),
    hInst, nil)
  discard SendMessage(inspector.btnClose, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.mainTree = CreateWindowExW(DWORD(WS_EX_CLIENTEDGE), WC_TREEVIEWW, nil,
    DWORD(treeStyle), 0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.mainTree, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.statusBar = CreateWindowExW(0, STATUSCLASSNAMEW, nil,
    WS_CHILD or WS_VISIBLE or SBARS_SIZEGRIP,
    0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.statusBar, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  discard SendMessage(inspector.statusBar, SB_SIMPLE, WPARAM(FALSE), 0)

  discard EnableWindow(inspector.btnInvoke, FALSE)
  discard EnableWindow(inspector.btnFocus, FALSE)
  discard EnableWindow(inspector.btnClose, FALSE)
  discard EnableWindow(inspector.btnHighlight, FALSE)
  discard EnableWindow(inspector.btnExpand, FALSE)

  updateStatusBar(inspector)
  rebuildElementTree(inspector)
  applyLayout(inspector)

proc updateHighlightColor(inspector: InspectorWindow; hexColor: string) =
  let parsed = parseColorRef(hexColor)
  if parsed.isSome:
    inspector.highlightColor = parsed.get()
  else:
    inspector.highlightColor = COLORREF(RGB(255, 0, 0))
  inspector.state.highlightColor = colorRefToHex(inspector.highlightColor)

proc updateStatusBar(inspector: InspectorWindow) =
  if inspector.statusBar == 0:
    return
  let version = if inspector.uiaVersion.len > 0: inspector.uiaVersion else: getUiaCoreVersion()
  let versionText = fmt"UIA version: {version}"
  let hintText = "Ctrl+C copies Acc path from the UIA Tree"
  discard SendMessage(inspector.statusBar, SB_SETTEXTW, 0, cast[LPARAM](newWideCString(versionText)))
  discard SendMessage(inspector.statusBar, SB_SETTEXTW, 1, cast[LPARAM](newWideCString(hintText)))

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
    var cc = default(TCHOOSECOLORW)
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
  of idRefresh:
    rebuildElementTree(inspector)
  else:
    discard

proc handleNotify(inspector: InspectorWindow; lParam: LPARAM) =
  let hdr = cast[ptr NMHDR](lParam)
  if hdr.hwndFrom == inspector.mainTree and hdr.code == UINT(TVN_SELCHANGEDW):
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
        for i in 0 ..< inspector.splitters.len:
          let r = inspector.splitters[i]
          if x >= r.left and x <= r.right and y >= r.top and y <= r.bottom:
            inspector.splitterDragging = i
            inspector.dragStartX = int(x)
            inspector.dragStartY = int(y)
            inspector.dragStartLeft = inspector.state.leftWidth
            inspector.dragStartMiddle = inspector.state.middleWidth
            inspector.dragStartProperties = inspector.state.propertiesHeight
            inspector.lastFocus = GetFocus()
            discard SetCapture(hwnd)
            break
      return 0
    of WM_MOUSEMOVE:
      if inspector != nil and inspector.splitterDragging >= 0:
        let x = lParamX(lParam)
        let y = lParamY(lParam)
        case inspector.splitterDragging
        of 0:
          inspector.state.leftWidth = inspector.dragStartLeft + (int(x) - inspector.dragStartX)
        of 1:
          inspector.state.middleWidth = inspector.dragStartMiddle + (int(x) - inspector.dragStartX)
        of 2:
          inspector.state.propertiesHeight = inspector.dragStartProperties + (int(y) - inspector.dragStartY)
        else:
          discard
        applyLayout(inspector)
      return 0
    of WM_LBUTTONUP:
      if inspector != nil and inspector.splitterDragging >= 0:
        inspector.splitterDragging = -1
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
        for r in inspector.splitters:
          discard FillRect(hdc, addr r, sashBrush)
        discard DeleteObject(sashBrush)
        discard EndPaint(hwnd, addr ps)
        return 0
    of WM_TIMER:
      if inspector != nil and UINT_PTR(wParam) == expandTimerId:
        handleExpandTimer(inspector)
      return 0
    of WM_SETCURSOR:
      if inspector != nil:
        var pt: POINT
        discard GetCursorPos(addr pt)
        discard ScreenToClient(hwnd, addr pt)
        for i, r in inspector.splitters:
          if pt.x >= r.left and pt.x <= r.right and pt.y >= r.top and pt.y <= r.bottom:
            let cursorId = if i == 2: IDC_SIZENS else: IDC_SIZEWE
            discard SetCursor(LoadCursorW(0, cursorId))
            return 1
      discard
    of WM_DESTROY:
      if inspector != nil:
        releaseNodes(inspector)
        saveInspectorState(inspector.statePath, inspector.state, inspector.logger)
        KillTimer(hwnd, UINT_PTR(expandTimerId))
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
  var icc = default(TINITCOMMONCONTROLSEX)
  icc.dwSize = DWORD(sizeof(icc))
  icc.dwICC = ICC_TREEVIEW_CLASSES or ICC_LISTVIEW_CLASSES or ICC_BAR_CLASSES
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
    statePath: if statePath.len > 0: statePath else: defaultInspectorStatePath(),
    splitterDragging: -1
  )
  insp.state = loadInspectorState(insp.statePath, logger)
  insp.uiaVersion = getUiaCoreVersion()
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

proc runInspectorApp*(statePath: string = ""): int =
  ## Standalone entry point for building the inspector as its own executable.
  var logger = newLogger()
  var uia: Uia
  try:
    uia = initUia()
  except CatchableError as exc:
    if logger != nil:
      logger.error("Failed to initialize UIA for inspector", [("error", exc.msg)])
    return 1

  defer:
    if not uia.isNil:
      uia.shutdown()

  let resolvedStatePath =
    if statePath.len > 0: statePath
    else: defaultInspectorStatePath()

  if not showInspectorWindow(uia, logger, resolvedStatePath):
    return 1

  var msg: MSG
  while true:
    if inspectors.len == 0:
      break
    let res = GetMessage(addr msg, 0, 0, 0)
    if res == 0 or res == -1:
      break
    discard TranslateMessage(addr msg)
    discard DispatchMessage(addr msg)

  0

when isMainModule:
  quit(runInspectorApp())
