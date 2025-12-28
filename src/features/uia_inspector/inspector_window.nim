when system.hostOS != "windows":
  {.error: "UIA inspector window is only supported on Windows.".}

import std/[options, os, sets, strformat, strutils, tables]

import winim/lean
import winim/com
import winim/inc/commctrl
import winim/inc/commdlg
import winim/inc/uiautomation
import winim/inc/winver
import winim/inc/psapi

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
  idUiaFilterEdit = 1007
  idPropertiesList = 1100
  idPatternsTree = 1101

  idMenuHighlightColor = 2001

type
  PropertyRow = object
    name: string
    value: string
    propertyId: PROPERTYID

  PatternAction = object
    patternId: PATTERNID
    action: string

type
  InspectorWindow = ref object
    hwnd: HWND
    gbWindowList: HWND
    gbWindowInfo: HWND
    gbProperties: HWND
    gbPatterns: HWND
    windowTitleLabel: HWND
    windowTitleValue: HWND
    windowHandleLabel: HWND
    windowHandleValue: HWND
    windowPosLabel: HWND
    windowPosValue: HWND
    windowClassInfoLabel: HWND
    windowClassValue: HWND
    windowProcessLabel: HWND
    windowProcessValue: HWND
    windowPidLabel: HWND
    windowPidValue: HWND
    filterVisible: HWND
    filterTitle: HWND
    filterActivate: HWND
    windowFilterLabel: HWND
    windowFilterEdit: HWND
    windowClassFilterLabel: HWND
    windowClassEdit: HWND
    btnRefresh: HWND
    windowList: HWND
    propertiesList: HWND
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
    expandTree: HWND
    btnInvoke: HWND
    btnFocus: HWND
    btnHighlight: HWND
    btnClose: HWND
    btnExpand: HWND
    uiaFilterLabel: HWND
    uiaFilterEdit: HWND
    uiaMaxDepth: int
    uiaFilterText: string
    propertyRows: seq[PropertyRow]
    patternActions: Table[HTREEITEM, PatternAction]
    patternCopyTexts: Table[HTREEITEM, string]
    accPath: string

var inspectors = initTable[HWND, InspectorWindow]()
var commonControlsReady = false
proc updateStatusBar(inspector: InspectorWindow)
proc resetWindowInfo(inspector: InspectorWindow)
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

proc safeLocalizedControlType(element: ptr IUIAutomationElement): string =
  try:
    var val: VARIANT
    let hr = element.GetCurrentPropertyValue(UIA_LocalizedControlTypePropertyId,
      addr val)
    if FAILED(hr) or val.vt != VT_BSTR or val.bstrVal.isNil:
      return ""
    $val.bstrVal
  except CatchableError:
    ""

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

proc readEditText(hwnd: HWND): string =
  let len = int(SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0))
  if len <= 0:
    return ""
  var buffer = newSeq[WCHAR](len + 1)
  discard GetWindowTextW(hwnd, addr buffer[0], cint(buffer.len))
  $cast[WideCString](addr buffer[0])

proc windowText(hwnd: HWND): string =
  var buffer = newSeq[WCHAR](256)
  let copied = GetWindowTextW(hwnd, addr buffer[0], cint(buffer.len))
  if copied <= 0:
    return ""
  $cast[WideCString](addr buffer[0])

proc windowClassName(hwnd: HWND): string =
  var buffer = newSeq[WCHAR](256)
  let copied = GetClassNameW(hwnd, addr buffer[0], cint(buffer.len))
  if copied <= 0:
    return ""
  $cast[WideCString](addr buffer[0])

proc copyToClipboard(text: string; logger: Logger = nil) =
  let sizeBytes = (text.len + 1) * int(sizeof(WCHAR))
  let hMem = GlobalAlloc(GMEM_MOVEABLE, SIZE_T(sizeBytes))
  if hMem == 0:
    if logger != nil:
      logger.error("Failed to allocate clipboard buffer")
    return

  let buffer = GlobalLock(hMem)
  if buffer.isNil:
    discard GlobalFree(hMem)
    if logger != nil:
      logger.error("Failed to lock clipboard buffer")
    return

  let wide = newWideCString(text)
  copyMem(buffer, unsafeAddr wide[0], sizeBytes)
  discard GlobalUnlock(hMem)

  if OpenClipboard(0) == 0:
    discard GlobalFree(hMem)
    if logger != nil:
      logger.error("Failed to open clipboard")
    return

  discard EmptyClipboard()
  if SetClipboardData(CF_UNICODETEXT, hMem) == 0:
    discard GlobalFree(hMem)
    if logger != nil:
      logger.error("Failed to set clipboard data")
  discard CloseClipboard()

proc processName(pid: DWORD): string =
  let handle = OpenProcess(PROCESS_QUERY_INFORMATION or PROCESS_VM_READ, FALSE, pid)
  if handle == 0:
    return ""
  defer: discard CloseHandle(handle)
  var buffer: array[MAX_PATH, WCHAR]
  let copied = GetModuleBaseNameW(handle, 0, addr buffer[0], DWORD(buffer.len))
  if copied == 0:
    return ""
  $cast[WideCString](addr buffer[0])

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
  let path = $cast[WideCString](addr pathBuf[0])
  let ver = fileVersion(path)
  if ver.len > 0:
    ver
  else:
    "Unknown"

proc nodeLabel(element: ptr IUIAutomationElement): string =
  let name = safeCurrentName(element)
  let automationId = safeAutomationId(element)
  let localized = safeLocalizedControlType(element)
  let ctrlType = if localized.len > 0: localized else: controlTypeName(safeControlType(element))

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

type
  ElementSubtree = ref object
    element: ptr IUIAutomationElement
    children: seq[ElementSubtree]
    matchesFilter: bool
    maxDepth: int

proc setTreeItemBold(tree: HWND; item: HTREEITEM; bold: bool) =
  var tvi: TVITEMW
  tvi.mask = UINT(TVIF_STATE)
  tvi.hItem = item
  tvi.stateMask = UINT(TVIS_BOLD)
  tvi.state = if bold: UINT(TVIS_BOLD) else: 0
  discard TreeView_SetItem(tree, addr tvi)

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

proc elementMatchesFilter(element: ptr IUIAutomationElement; filterLower: string): bool =
  if element.isNil or filterLower.len == 0:
    return false

  let localized = safeLocalizedControlType(element).toLower()
  let ctrlTypeFallback = controlTypeName(safeControlType(element)).toLower()
  let name = safeCurrentName(element).toLower()
  let automationId = safeAutomationId(element).toLower()

  for candidate in [localized, ctrlTypeFallback, name, automationId]:
    if candidate.len > 0 and candidate.find(filterLower) >= 0:
      return true
  false

proc collectSubtree(inspector: InspectorWindow; walker: ptr IUIAutomationTreeWalker;
    element: ptr IUIAutomationElement; depth: int; filterLower: string;
    hasFilter: bool; forceInclude: bool): ElementSubtree =
  if element.isNil:
    return nil

  var maxDepth = depth
  var children: seq[ElementSubtree] = @[]
  var child: ptr IUIAutomationElement
  let hrFirst = walker.GetFirstChildElement(element, addr child)
  if SUCCEEDED(hrFirst) and not child.isNil:
    var current = child
    while current != nil:
      let childNode = collectSubtree(inspector, walker, current, depth + 1,
        filterLower, hasFilter, false)

      var next: ptr IUIAutomationElement
      let hrNext = walker.GetNextSiblingElement(current, addr next)
      if not childNode.isNil:
        maxDepth = max(maxDepth, childNode.maxDepth)
        children.add(childNode)
      else:
        discard current.Release()
      if FAILED(hrNext) or hrNext == S_FALSE:
        break
      current = next

  let matchesSelf = hasFilter and elementMatchesFilter(element, filterLower)
  let includeNode = forceInclude or (not hasFilter) or matchesSelf or children.len > 0
  if not includeNode:
    return nil

  result = ElementSubtree(
    element: element,
    children: children,
    matchesFilter: matchesSelf,
    maxDepth: maxDepth
  )

proc renderSubtree(inspector: InspectorWindow; tree: HWND; node: ElementSubtree;
    parent: HTREEITEM) =
  if node.isNil:
    return
  let item = addTreeItem(tree, parent, nodeLabel(node.element),
    cast[LPARAM](node.element))
  inspector.nodes[item] = node.element
  setTreeItemBold(tree, item, node.matchesFilter)
  for child in node.children:
    renderSubtree(inspector, tree, child, item)

proc rebuildElementTree(inspector: InspectorWindow) =
  TreeView_DeleteAllItems(inspector.mainTree)
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

  inspector.uiaFilterText =
    if inspector.uiaFilterEdit != 0: readEditText(inspector.uiaFilterEdit).strip()
    else: ""
  let filterLower = inspector.uiaFilterText.toLower()
  let hasFilter = filterLower.len > 0

  let rootNode = collectSubtree(inspector, walker, root, 0, filterLower, hasFilter, true)
  if rootNode.isNil:
    if inspector.logger != nil:
      inspector.logger.warn("UIA tree filtered to zero nodes")
    inspector.uiaMaxDepth = 0
    populateProperties(inspector, nil)
    updateStatusBar(inspector)
    return

  inspector.uiaMaxDepth = max(1, rootNode.maxDepth + 1)
  renderSubtree(inspector, inspector.mainTree, rootNode, TVI_ROOT)

  let rootItem = TreeView_GetRoot(inspector.mainTree)
  TreeView_Expand(inspector.mainTree, rootItem, UINT(TVE_EXPAND))
  discard TreeView_SelectItem(inspector.mainTree, rootItem)
  populateProperties(inspector, root)
  updateStatusBar(inspector)

type WindowEnumContext = object
  inspector: InspectorWindow
  requireVisible: bool
  requireTitle: bool
  titleFilter: string
  classFilter: string
  index: int

proc syncFilterState(inspector: InspectorWindow) =
  if inspector.filterVisible != 0:
    inspector.state.filterVisible =
      SendMessage(inspector.filterVisible, BM_GETCHECK, 0, 0) == BST_CHECKED
  if inspector.filterTitle != 0:
    inspector.state.filterTitle =
      SendMessage(inspector.filterTitle, BM_GETCHECK, 0, 0) == BST_CHECKED
  if inspector.filterActivate != 0:
    inspector.state.filterActivate =
      SendMessage(inspector.filterActivate, BM_GETCHECK, 0, 0) == BST_CHECKED

proc selectedWindowHandle(inspector: InspectorWindow): HWND =
  let idx = int(SendMessage(inspector.windowList, LVM_GETNEXTITEM, WPARAM(-1),
    LPARAM(LVNI_SELECTED)))
  if idx < 0:
    return HWND(0)
  var item: LVITEMW
  item.mask = LVIF_PARAM
  item.iItem = idx.cint
  discard SendMessage(inspector.windowList, LVM_GETITEMW, 0, cast[LPARAM](addr item))
  cast[HWND](item.lParam)

proc refreshWindowList(inspector: InspectorWindow) =
  if inspector.windowList == 0:
    return
  discard SendMessage(inspector.windowList, LVM_DELETEALLITEMS, 0, 0)

  syncFilterState(inspector)
  let filterVisible = inspector.state.filterVisible
  let filterTitle = inspector.state.filterTitle

  var ctx = WindowEnumContext(
    inspector: inspector,
    requireVisible: filterVisible,
    requireTitle: filterTitle,
    titleFilter: readEditText(inspector.windowFilterEdit).strip().toLower(),
    classFilter: readEditText(inspector.windowClassEdit).strip().toLower(),
    index: 0
  )

  EnumWindows(proc(hwnd: HWND; lParam: LPARAM): BOOL {.stdcall.} =
    let ctxPtr = cast[ptr WindowEnumContext](lParam)
    if ctxPtr == nil or ctxPtr.inspector.isNil:
      return TRUE
    let insp = ctxPtr.inspector
    if hwnd == insp.hwnd:
      return TRUE
    if ctxPtr.requireVisible and IsWindowVisible(hwnd) == 0:
      return TRUE

    let title = windowText(hwnd)
    if ctxPtr.requireTitle and title.len == 0:
      return TRUE
    let titleLower = title.toLower()
    if ctxPtr.titleFilter.len > 0 and titleLower.find(ctxPtr.titleFilter) < 0:
      return TRUE

    let cls = windowClassName(hwnd)
    let clsLower = cls.toLower()
    if ctxPtr.classFilter.len > 0 and clsLower.find(ctxPtr.classFilter) < 0:
      return TRUE

    var pid: DWORD
    discard GetWindowThreadProcessId(hwnd, addr pid)
    let procTextRaw = processName(pid)
    let procText = if procTextRaw.len > 0: procTextRaw else: fmt"PID {pid}"
    let pidText = $pid

    var item: LVITEMW
    item.mask = LVIF_TEXT or LVIF_PARAM
    item.iItem = ctxPtr.index.cint
    item.pszText = newWideCString(title)
    item.lParam = cast[LPARAM](hwnd)
    let insertIndex = SendMessage(insp.windowList, LVM_INSERTITEMW, 0,
      cast[LPARAM](addr item))
    if insertIndex != -1:
      var sub: LVITEMW
      sub.iItem = int32(insertIndex)
      sub.mask = LVIF_TEXT
      sub.iSubItem = 1
      sub.pszText = newWideCString(procText)
      discard SendMessage(insp.windowList, LVM_SETITEMTEXTW, 0, cast[LPARAM](addr sub))
      sub.iSubItem = 2
      sub.pszText = newWideCString(pidText)
      discard SendMessage(insp.windowList, LVM_SETITEMTEXTW, 0, cast[LPARAM](addr sub))
      inc ctxPtr.index
    TRUE
  , cast[LPARAM](addr ctx))

  saveInspectorState(inspector.statePath, inspector.state, inspector.logger)

proc handleWindowSelectionChanged(inspector: InspectorWindow) =
  syncFilterState(inspector)
  let hwndSel = inspector.selectedWindowHandle()
  if hwndSel == HWND(0):
    resetWindowInfo(inspector)
    return
  if SendMessage(inspector.filterActivate, BM_GETCHECK, 0, 0) == BST_CHECKED:
    discard SetForegroundWindow(hwndSel)
    discard SetFocus(hwndSel)

  if inspector.uia.isNil:
    resetWindowInfo(inspector)
    return

  try:
    let element = inspector.uia.fromWindowHandle(hwndSel)
    inspector.uia.setRootElement(element)
    rebuildElementTree(inspector)
  except CatchableError as exc:
    if inspector.logger != nil:
      inspector.logger.error("Failed to build inspector tree from window",
        [("error", exc.msg)])
    inspector.uia.setRootElement(nil)
  resetWindowInfo(inspector)

proc setEditText(hwnd: HWND; value: string) =
  discard SetWindowTextW(hwnd, newWideCString(value))

proc resetWindowInfo(inspector: InspectorWindow) =
  setEditText(inspector.windowTitleValue, "No selection")
  for hwnd in [inspector.windowHandleValue, inspector.windowPosValue,
      inspector.windowClassValue, inspector.windowProcessValue, inspector.windowPidValue]:
    if hwnd != 0:
      setEditText(hwnd, "")

proc boolToStr(flag: bool): string =
  if flag: "True" else: "False"

proc roleText(role: LONG): string =
  var buf = newSeq[WCHAR](128)
  let written = GetRoleTextW(role, addr buf[0], cint(buf.len))
  if written > 0:
    buf.setLen(written)
    $cast[WideCString](addr buf[0])
  else:
    fmt"Role({role})"

proc propertyValueString(element: ptr IUIAutomationElement; propertyId: PROPERTYID): string =
  if element.isNil:
    return ""
  if propertyId == UIA_ControlTypePropertyId:
    return controlTypeName(safeControlType(element))
  if propertyId == UIA_LocalizedControlTypePropertyId:
    return safeLocalizedControlType(element)
  if propertyId == UIA_BoundingRectanglePropertyId:
    let bounds = safeBoundingRect(element)
    if bounds.isSome:
      let (l, t, w, h) = bounds.get()
      return fmt"({l.int},{t.int}) {w.int}x{h.int}"
    else:
      return "Unavailable"

  var val: VARIANT
  let hr = element.GetCurrentPropertyValue(propertyId, addr val)
  defer:
    discard VariantClear(addr val)
  if FAILED(hr):
    return fmt"Error 0x{hr:X}"
  case val.vt
  of VT_EMPTY, VT_NULL:
    ""
  of VT_BSTR:
    if val.bstrVal.isNil: "" else: $val.bstrVal
  of VT_I1, VT_I2, VT_I4, VT_UI1, VT_UI2, VT_UI4:
    $val.lVal
  of VT_I8, VT_UI8:
    $val.llVal
  of VT_BOOL:
    if val.boolVal == VARIANT_TRUE: "True" else: "False"
  of VT_UNKNOWN:
    "Object"
  else:
    fmt"VT({val.vt})"

proc populatePropertyList(inspector: InspectorWindow; element: ptr IUIAutomationElement) =
  inspector.propertyRows.setLen(0)
  discard SendMessage(inspector.propertiesList, LVM_DELETEALLITEMS, 0, 0)

  if element.isNil:
    return

  let propertyIds: seq[(string, PROPERTYID)] = @[
    ("ControlType", UIA_ControlTypePropertyId),
    ("LocalizedControlType", UIA_LocalizedControlTypePropertyId),
    ("Name", UIA_NamePropertyId),
    ("Value", UIA_ValueValuePropertyId),
    ("AutomationId", UIA_AutomationIdPropertyId),
    ("BoundingRectangle", UIA_BoundingRectanglePropertyId),
    ("ClassName", UIA_ClassNamePropertyId),
    ("HelpText", UIA_HelpTextPropertyId),
    ("AccessKey", UIA_AccessKeyPropertyId),
    ("AcceleratorKey", UIA_AcceleratorKeyPropertyId),
    ("HasKeyboardFocus", UIA_HasKeyboardFocusPropertyId),
    ("IsKeyboardFocusable", UIA_IsKeyboardFocusablePropertyId),
    ("ItemType", UIA_ItemTypePropertyId),
    ("ProcessId", UIA_ProcessIdPropertyId),
    ("IsEnabled", UIA_IsEnabledPropertyId),
    ("IsPassword", UIA_IsPasswordPropertyId),
    ("IsOffscreen", UIA_IsOffscreenPropertyId),
    ("FrameworkId", UIA_FrameworkIdPropertyId),
    ("IsRequiredForForm", UIA_IsRequiredForFormPropertyId),
    ("ItemStatus", UIA_ItemStatusPropertyId),
    ("LabeledBy", UIA_LabeledByPropertyId)
  ]

  for prop in propertyIds:
    let value = propertyValueString(element, prop[1])
    inspector.propertyRows.add(PropertyRow(name: prop[0], value: value, propertyId: prop[1]))

  for idx, row in inspector.propertyRows:
    var item: LVITEMW
    item.mask = LVIF_TEXT or LVIF_PARAM
    item.iItem = idx.cint
    item.pszText = newWideCString(row.name)
    item.lParam = LPARAM(idx)
    let inserted = SendMessage(inspector.propertiesList, LVM_INSERTITEMW, 0,
      cast[LPARAM](addr item))
    if inserted != -1:
      var sub: LVITEMW
      sub.mask = LVIF_TEXT
      sub.iItem = int32(inserted)
      sub.iSubItem = 1
      sub.pszText = newWideCString(row.value)
      discard SendMessage(inspector.propertiesList, LVM_SETITEMTEXTW, 0,
        cast[LPARAM](addr sub))

proc populateWindowInfo(inspector: InspectorWindow; element: ptr IUIAutomationElement) =
  if element.isNil:
    resetWindowInfo(inspector)
    return
  setEditText(inspector.windowTitleValue, safeCurrentName(element))
  let className = safeClassName(element)
  setEditText(inspector.windowClassValue, className)
  let automationId = safeAutomationId(element)
  let nativeHwnd = nativeWindowHandle(element)
  if nativeHwnd != 0:
    setEditText(inspector.windowHandleValue, fmt"0x{cast[uint](nativeHwnd):X}")
    var rect: RECT
    if GetWindowRect(HWND(nativeHwnd), addr rect) != 0:
      let width = rect.right - rect.left
      let height = rect.bottom - rect.top
      setEditText(inspector.windowPosValue,
        fmt"({rect.left},{rect.top}) {width}x{height}")
  else:
    setEditText(inspector.windowHandleValue, "")
    setEditText(inspector.windowPosValue, "Unavailable")
  if automationId.len > 0 and className.len == 0:
    setEditText(inspector.windowClassValue, automationId)

  var pidVar: VARIANT
  let hr = element.GetCurrentPropertyValue(UIA_ProcessIdPropertyId, addr pidVar)
  if SUCCEEDED(hr) and (pidVar.vt == VT_I4 or pidVar.vt == VT_INT):
    let pid = DWORD(pidVar.lVal)
    setEditText(inspector.windowPidValue, $pid)
    setEditText(inspector.windowProcessValue, processName(pid))
  else:
    setEditText(inspector.windowPidValue, "")
    setEditText(inspector.windowProcessValue, "")
  discard VariantClear(addr pidVar)

proc addPatternNode(inspector: InspectorWindow; parent: HTREEITEM; text: string;
    copyText: string; action: PatternAction = PatternAction()): HTREEITEM =
  let item = addTreeItem(inspector.patternsTree, parent, text)
  if copyText.len > 0:
    inspector.patternCopyTexts[item] = copyText
  if action.action.len > 0:
    inspector.patternActions[item] = action
  item

proc getPattern[T](element: ptr IUIAutomationElement; patternId: PATTERNID;
    resultPtr: var ptr T): bool =
  resultPtr = nil
  if element.isNil:
    return false
  let hr = element.GetCurrentPattern(patternId,
    cast[ptr ptr IUnknown](addr resultPtr))
  if FAILED(hr) or resultPtr.isNil:
    return false
  true

proc addBoolChild(inspector: InspectorWindow; parent: HTREEITEM; label: string; flag: bool) =
  discard addPatternNode(inspector, parent, fmt"{label}: {boolToStr(flag)}",
    fmt"{label}: {boolToStr(flag)}")

proc populatePatterns(inspector: InspectorWindow; element: ptr IUIAutomationElement) =
  TreeView_DeleteAllItems(inspector.patternsTree)
  inspector.patternActions.clear()
  inspector.patternCopyTexts.clear()

  if element.isNil:
    discard addTreeItem(inspector.patternsTree, TVI_ROOT, "No patterns available")
    return

  var anyPattern = false

  var invoke: ptr IUIAutomationInvokePattern
  if getPattern(element, UIA_InvokePatternId, invoke):
    defer: discard invoke.Release()
    let root = addPatternNode(inspector, TVI_ROOT, "Invoke", "Invoke")
    discard addPatternNode(inspector, root, "Action: Invoke", "Invoke",
      PatternAction(patternId: UIA_InvokePatternId, action: "Invoke"))
    discard TreeView_Expand(inspector.patternsTree, root, UINT(TVE_EXPAND))
    anyPattern = true

  var legacy: ptr IUIAutomationLegacyIAccessiblePattern
  if getPattern(element, UIA_LegacyIAccessiblePatternId, legacy):
    defer: discard legacy.Release()
    var name: BSTR
    var value: BSTR
    var desc: BSTR
    var role: LONG
    discard legacy.get_CurrentName(addr name)
    discard legacy.get_CurrentValue(addr value)
    discard legacy.get_CurrentDescription(addr desc)
    discard legacy.get_CurrentRole(addr role)
    let root = addPatternNode(inspector, TVI_ROOT, "LegacyIAccessible", "LegacyIAccessible")
    discard addPatternNode(inspector, root, "CurrentName: " & (if name.isNil: "" else: $name),
      if name.isNil: "" else: $name)
    discard addPatternNode(inspector, root, "CurrentValue: " & (if value.isNil: "" else: $value),
      if value.isNil: "" else: $value)
    discard addPatternNode(inspector, root, "CurrentDescription: " & (if desc.isNil: "" else: $desc),
      if desc.isNil: "" else: $desc)
    discard addPatternNode(inspector, root, "CurrentRole: " & roleText(role),
      roleText(role))
    discard addPatternNode(inspector, root, "Action: DoDefaultAction", "DoDefaultAction",
      PatternAction(patternId: UIA_LegacyIAccessiblePatternId, action: "DoDefaultAction"))
    if not name.isNil: SysFreeString(name)
    if not value.isNil: SysFreeString(value)
    if not desc.isNil: SysFreeString(desc)
    discard TreeView_Expand(inspector.patternsTree, root, UINT(TVE_EXPAND))
    anyPattern = true

  var selectionItem: ptr IUIAutomationSelectionItemPattern
  if getPattern(element, UIA_SelectionItemPatternId, selectionItem):
    defer: discard selectionItem.Release()
    var isSelected: BOOL
    discard selectionItem.get_CurrentIsSelected(addr isSelected)
    let root = addPatternNode(inspector, TVI_ROOT, "SelectionItem", "SelectionItem")
    addBoolChild(inspector, root, "CurrentIsSelected", isSelected != 0)
    discard addPatternNode(inspector, root, "Action: Select", "Select",
      PatternAction(patternId: UIA_SelectionItemPatternId, action: "Select"))
    discard TreeView_Expand(inspector.patternsTree, root, UINT(TVE_EXPAND))
    anyPattern = true

  var valuePattern: ptr IUIAutomationValuePattern
  if getPattern(element, UIA_ValuePatternId, valuePattern):
    defer: discard valuePattern.Release()
    var current: BSTR
    var readOnly: BOOL
    discard valuePattern.get_CurrentValue(addr current)
    discard valuePattern.get_CurrentIsReadOnly(addr readOnly)
    let root = addPatternNode(inspector, TVI_ROOT, "Value", "Value")
    discard addPatternNode(inspector, root, "CurrentValue: " & (if current.isNil: "" else: $current),
      if current.isNil: "" else: $current)
    addBoolChild(inspector, root, "CurrentIsReadOnly", readOnly != 0)
    discard addPatternNode(inspector, root, "Action: SetValue (uses clipboard text)", "SetValue",
      PatternAction(patternId: UIA_ValuePatternId, action: "SetValue"))
    if not current.isNil: SysFreeString(current)
    discard TreeView_Expand(inspector.patternsTree, root, UINT(TVE_EXPAND))
    anyPattern = true

  if not anyPattern:
    discard addTreeItem(inspector.patternsTree, TVI_ROOT, "No patterns available")

proc readClipboardText(): Option[string] =
  if OpenClipboard(0) == 0:
    return
  defer: discard CloseClipboard()
  let handle = GetClipboardData(CF_UNICODETEXT)
  if handle == 0:
    return
  let data = GlobalLock(handle)
  if data.isNil:
    return
  defer: discard GlobalUnlock(handle)
  let text = $cast[WideCString](data)
  if text.len > 0:
    result = some(text)

proc executePatternAction(inspector: InspectorWindow; element: ptr IUIAutomationElement;
    action: PatternAction) =
  if element.isNil:
    return
  case action.action
  of "Invoke":
    var pattern: ptr IUIAutomationInvokePattern
    if getPattern(element, UIA_InvokePatternId, pattern):
      defer: discard pattern.Release()
      discard pattern.Invoke()
  of "DoDefaultAction":
    var pattern: ptr IUIAutomationLegacyIAccessiblePattern
    if getPattern(element, UIA_LegacyIAccessiblePatternId, pattern):
      defer: discard pattern.Release()
      discard pattern.DoDefaultAction()
  of "Select":
    var pattern: ptr IUIAutomationSelectionItemPattern
    if getPattern(element, UIA_SelectionItemPatternId, pattern):
      defer: discard pattern.Release()
      discard pattern.Select()
  of "SetValue":
    var pattern: ptr IUIAutomationValuePattern
    if getPattern(element, UIA_ValuePatternId, pattern):
      defer: discard pattern.Release()
      let clip = readClipboardText()
      if clip.isSome:
        let text = clip.get()
        let wide = newWideCString(text)
        let bstr = SysAllocString(cast[LPCWSTR](wide))
        if bstr != nil:
          discard pattern.SetValue(bstr)
          SysFreeString(bstr)
  else:
    discard

proc computeAccPath(element: ptr IUIAutomationElement): Option[string] =
  ## Legacy accessibility path traversal is unavailable with the current bindings.
  ## Return none to fall back gracefully.
  discard element

proc populateProperties(inspector: InspectorWindow; element: ptr IUIAutomationElement) =
  TreeView_DeleteAllItems(inspector.patternsTree)
  inspector.accPath = ""
  if element.isNil:
    discard EnableWindow(inspector.btnInvoke, FALSE)
    discard EnableWindow(inspector.btnFocus, FALSE)
    discard EnableWindow(inspector.btnClose, FALSE)
    discard EnableWindow(inspector.btnHighlight, FALSE)
    discard EnableWindow(inspector.btnExpand, FALSE)
    populatePropertyList(inspector, nil)
    populatePatterns(inspector, nil)
    resetWindowInfo(inspector)
    updateStatusBar(inspector)
    return

  discard EnableWindow(inspector.btnInvoke, TRUE)
  discard EnableWindow(inspector.btnFocus, TRUE)
  discard EnableWindow(inspector.btnClose, TRUE)
  discard EnableWindow(inspector.btnHighlight, TRUE)
  discard EnableWindow(inspector.btnExpand, TRUE)

  populateWindowInfo(inspector, element)
  populatePropertyList(inspector, element)
  populatePatterns(inspector, element)
  let accOpt = computeAccPath(element)
  if accOpt.isSome:
    inspector.accPath = accOpt.get()
  else:
    inspector.accPath = ""
  updateStatusBar(inspector)

proc currentSelection(inspector: InspectorWindow): ptr IUIAutomationElement =
  let selected = TreeView_GetSelection(inspector.mainTree)
  if selected == 0 or selected notin inspector.nodes:
    return nil
  inspector.nodes[selected]

proc expandTarget(inspector: InspectorWindow): (HWND, HTREEITEM) =
  var focused = GetFocus()
  if focused == inspector.mainTree:
    let sel = TreeView_GetSelection(inspector.mainTree)
    if sel != 0:
      return (inspector.mainTree, sel)

  let mainSel = TreeView_GetSelection(inspector.mainTree)
  if mainSel != 0:
    return (inspector.mainTree, mainSel)

  (inspector.mainTree, TreeView_GetRoot(inspector.mainTree))

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

proc collectExpandQueue(inspector: InspectorWindow; tree: HWND; start: HTREEITEM) =
  inspector.expandQueue.setLen(0)
  if tree == 0 or start == 0:
    return
  inspector.expandQueue.add(start)

proc beginExpandAll(inspector: InspectorWindow; tree: HWND; start: HTREEITEM) =
  if inspector.expandActive:
    return
  inspector.expandTree = tree
  collectExpandQueue(inspector, tree, start)
  if inspector.expandQueue.len == 0 or inspector.expandTree == 0:
    if inspector.logger != nil:
      inspector.logger.warn("Expand-all requested with no tree selection")
    inspector.expandTree = 0
    return
  inspector.expandActive = true
  discard EnableWindow(inspector.btnExpand, FALSE)
  discard SetTimer(inspector.hwnd, UINT_PTR(expandTimerId), UINT(10), nil)

proc handleExpandTimer(inspector: InspectorWindow) =
  var processed = 0
  while inspector.expandQueue.len > 0 and processed < 64:
    let item = inspector.expandQueue.pop()
    if inspector.expandTree == 0:
      break
    discard TreeView_Expand(inspector.expandTree, item, UINT(TVE_EXPAND))
    var child = TreeView_GetChild(inspector.expandTree, item)
    while child != 0:
      inspector.expandQueue.add(child)
      child = TreeView_GetNextSibling(inspector.expandTree, child)
    inc processed

  if inspector.expandQueue.len == 0:
    KillTimer(inspector.hwnd, UINT_PTR(expandTimerId))
    inspector.expandActive = false
    inspector.expandTree = 0
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
  MoveWindow(inspector.filterVisible, groupInnerLeft.cint, currentY.cint, 80, 20, TRUE)
  MoveWindow(inspector.filterTitle, (groupInnerLeft + 90).cint, currentY.cint, 80, 20, TRUE)
  MoveWindow(inspector.filterActivate, (groupInnerLeft + 180).cint, currentY.cint, 100, 20, TRUE)
  currentY += 24
  MoveWindow(inspector.windowClassFilterLabel, groupInnerLeft.cint, currentY.cint,
    (leftWidth - 2 * groupPadding).int32, 16, TRUE)
  currentY += 18
  MoveWindow(inspector.windowClassEdit, groupInnerLeft.cint, currentY.cint,
    (leftWidth - 2 * groupPadding).int32, 22, TRUE)
  currentY += 28
  MoveWindow(inspector.btnRefresh, groupInnerLeft.cint, currentY.cint, 140, buttonHeight.int32, TRUE)
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
  var infoY = contentTop + groupPadding
  let labelWidth = 80
  let valueWidth = max(infoInnerWidth - labelWidth - 6, 80)
  let rowHeight = 18

  MoveWindow(inspector.windowTitleLabel, (middleX + groupPadding).cint, infoY.cint,
    labelWidth.int32, rowHeight.int32, TRUE)
  MoveWindow(inspector.windowTitleValue, (middleX + groupPadding + labelWidth + 4).cint,
    infoY.cint, valueWidth.int32, rowHeight.int32, TRUE)
  infoY += rowHeight + 4

  MoveWindow(inspector.windowHandleLabel, (middleX + groupPadding).cint, infoY.cint,
    labelWidth.int32, rowHeight.int32, TRUE)
  MoveWindow(inspector.windowHandleValue, (middleX + groupPadding + labelWidth + 4).cint,
    infoY.cint, valueWidth.int32, rowHeight.int32, TRUE)
  infoY += rowHeight + 4

  MoveWindow(inspector.windowPosLabel, (middleX + groupPadding).cint, infoY.cint,
    labelWidth.int32, rowHeight.int32, TRUE)
  MoveWindow(inspector.windowPosValue, (middleX + groupPadding + labelWidth + 4).cint,
    infoY.cint, valueWidth.int32, rowHeight.int32, TRUE)
  infoY += rowHeight + 4

  MoveWindow(inspector.windowClassInfoLabel, (middleX + groupPadding).cint, infoY.cint,
    labelWidth.int32, rowHeight.int32, TRUE)
  MoveWindow(inspector.windowClassValue, (middleX + groupPadding + labelWidth + 4).cint,
    infoY.cint, valueWidth.int32, rowHeight.int32, TRUE)
  infoY += rowHeight + 4

  MoveWindow(inspector.windowProcessLabel, (middleX + groupPadding).cint, infoY.cint,
    labelWidth.int32, rowHeight.int32, TRUE)
  MoveWindow(inspector.windowProcessValue, (middleX + groupPadding + labelWidth + 4).cint,
    infoY.cint, valueWidth.int32, rowHeight.int32, TRUE)
  infoY += rowHeight + 4

  MoveWindow(inspector.windowPidLabel, (middleX + groupPadding).cint, infoY.cint,
    labelWidth.int32, rowHeight.int32, TRUE)
  MoveWindow(inspector.windowPidValue, (middleX + groupPadding + labelWidth + 4).cint,
    infoY.cint, valueWidth.int32, rowHeight.int32, TRUE)

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
  MoveWindow(inspector.propertiesList, (middleX + groupPadding).cint,
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

  let filterLabelHeight = 16
  let filterEditHeight = 22
  let filterTop = contentTop + groupPadding
  MoveWindow(inspector.uiaFilterLabel, (rightX + groupPadding).cint, filterTop.cint,
    (rightWidth - 2 * groupPadding).int32, filterLabelHeight.int32, TRUE)
  MoveWindow(inspector.uiaFilterEdit, (rightX + groupPadding).cint,
    (filterTop + filterLabelHeight + 4).cint,
    (rightWidth - 2 * groupPadding).int32, filterEditHeight.int32, TRUE)

  let treeTop = filterTop + filterLabelHeight + filterEditHeight + 8
  let treeHeight = max(usableHeight - (treeTop - contentTop), 40)
  MoveWindow(inspector.mainTree, rightX.cint, treeTop.cint, rightWidth.int32,
    treeHeight.int32, TRUE)

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
  inspector.patternActions = initTable[HTREEITEM, PatternAction]()
  inspector.patternCopyTexts = initTable[HTREEITEM, string]()
  inspector.propertyRows.setLen(0)
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

  inspector.filterVisible = CreateWindowExW(0, WC_BUTTON,
    newWideCString("Visible"), WS_CHILD or WS_VISIBLE or WS_TABSTOP or BS_AUTOCHECKBOX,
    0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.filterVisible, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  discard SendMessage(inspector.filterVisible, BM_SETCHECK,
    WPARAM(if inspector.state.filterVisible: BST_CHECKED else: BST_UNCHECKED), 0)

  inspector.filterTitle = CreateWindowExW(0, WC_BUTTON,
    newWideCString("Title"), WS_CHILD or WS_VISIBLE or WS_TABSTOP or BS_AUTOCHECKBOX,
    0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.filterTitle, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  discard SendMessage(inspector.filterTitle, BM_SETCHECK,
    WPARAM(if inspector.state.filterTitle: BST_CHECKED else: BST_UNCHECKED), 0)

  inspector.filterActivate = CreateWindowExW(0, WC_BUTTON,
    newWideCString("Activate"), WS_CHILD or WS_VISIBLE or WS_TABSTOP or BS_AUTOCHECKBOX,
    0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.filterActivate, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  discard SendMessage(inspector.filterActivate, BM_SETCHECK,
    WPARAM(if inspector.state.filterActivate: BST_CHECKED else: BST_UNCHECKED), 0)

  inspector.windowClassFilterLabel = CreateWindowExW(0, WC_STATIC,
    newWideCString("Class filter:"), WS_CHILD or WS_VISIBLE,
    0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.windowClassFilterLabel, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowClassEdit = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or WS_TABSTOP or ES_AUTOHSCROLL,
    0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.windowClassEdit, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.btnRefresh = CreateWindowExW(0, WC_BUTTON, newWideCString("Refresh window list"),
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
  col.pszText = newWideCString("Title")
  discard SendMessage(inspector.windowList, LVM_INSERTCOLUMNW, 0, cast[LPARAM](addr col))
  col.cx = 140
  col.pszText = newWideCString("Process")
  discard SendMessage(inspector.windowList, LVM_INSERTCOLUMNW, 1, cast[LPARAM](addr col))
  col.cx = 80
  col.pszText = newWideCString("ID")
  discard SendMessage(inspector.windowList, LVM_INSERTCOLUMNW, 2, cast[LPARAM](addr col))

  inspector.gbWindowInfo = CreateWindowExW(0, WC_BUTTON, newWideCString("Window Info"),
    WS_CHILD or WS_VISIBLE or BS_GROUPBOX, 0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.gbWindowInfo, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowTitleLabel = CreateWindowExW(0, WC_STATIC, newWideCString("Title:"),
    WS_CHILD or WS_VISIBLE, 0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.windowTitleLabel, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  inspector.windowTitleValue = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or ES_AUTOHSCROLL or ES_READONLY,
    0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.windowTitleValue, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowHandleLabel = CreateWindowExW(0, WC_STATIC, newWideCString("HWND:"),
    WS_CHILD or WS_VISIBLE, 0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.windowHandleLabel, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  inspector.windowHandleValue = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or ES_AUTOHSCROLL or ES_READONLY,
    0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.windowHandleValue, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowPosLabel = CreateWindowExW(0, WC_STATIC, newWideCString("Position:"),
    WS_CHILD or WS_VISIBLE, 0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.windowPosLabel, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  inspector.windowPosValue = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or ES_AUTOHSCROLL or ES_READONLY,
    0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.windowPosValue, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowClassInfoLabel = CreateWindowExW(0, WC_STATIC, newWideCString("Class:"),
    WS_CHILD or WS_VISIBLE, 0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.windowClassInfoLabel, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  inspector.windowClassValue = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or ES_AUTOHSCROLL or ES_READONLY,
    0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.windowClassValue, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowProcessLabel = CreateWindowExW(0, WC_STATIC, newWideCString("Process:"),
    WS_CHILD or WS_VISIBLE, 0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.windowProcessLabel, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  inspector.windowProcessValue = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or ES_AUTOHSCROLL or ES_READONLY,
    0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.windowProcessValue, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowPidLabel = CreateWindowExW(0, WC_STATIC, newWideCString("PID:"),
    WS_CHILD or WS_VISIBLE, 0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.windowPidLabel, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  inspector.windowPidValue = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or ES_AUTOHSCROLL or ES_READONLY,
    0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.windowPidValue, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.gbProperties = CreateWindowExW(0, WC_BUTTON, newWideCString("Properties"),
    WS_CHILD or WS_VISIBLE or BS_GROUPBOX, 0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.gbProperties, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  var propListStyle = WS_CHILD or WS_VISIBLE or WS_TABSTOP or LVS_REPORT or LVS_SHOWSELALWAYS or
      LVS_SINGLESEL or WS_BORDER
  inspector.propertiesList = CreateWindowExW(DWORD(WS_EX_CLIENTEDGE), WC_LISTVIEWW, nil,
    DWORD(propListStyle), 0, 0, 0, 0, inspector.hwnd, HMENU(idPropertiesList), hInst, nil)
  discard SendMessage(inspector.propertiesList, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  discard SendMessage(inspector.propertiesList, LVM_SETEXTENDEDLISTVIEWSTYLE, 0,
    LPARAM(LVS_EX_FULLROWSELECT or LVS_EX_GRIDLINES))
  var propCol: LVCOLUMNW
  propCol.mask = LVCF_TEXT or LVCF_WIDTH
  propCol.cx = 170
  propCol.pszText = newWideCString("Property")
  discard SendMessage(inspector.propertiesList, LVM_INSERTCOLUMNW, 0, cast[LPARAM](addr propCol))
  propCol.cx = 360
  propCol.pszText = newWideCString("Value")
  discard SendMessage(inspector.propertiesList, LVM_INSERTCOLUMNW, 1, cast[LPARAM](addr propCol))

  inspector.gbPatterns = CreateWindowExW(0, WC_BUTTON, newWideCString("Patterns"),
    WS_CHILD or WS_VISIBLE or BS_GROUPBOX, 0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.gbPatterns, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  var treeStyle = WS_CHILD or WS_VISIBLE or WS_TABSTOP or TVS_HASBUTTONS or
      TVS_LINESATROOT or TVS_HASLINES or WS_BORDER
  inspector.patternsTree = CreateWindowExW(DWORD(WS_EX_CLIENTEDGE), WC_TREEVIEWW, nil,
    DWORD(treeStyle), 0, 0, 0, 0, inspector.hwnd, HMENU(idPatternsTree), hInst, nil)
  discard SendMessage(inspector.patternsTree, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.btnHighlight = CreateWindowExW(0, WC_BUTTON,
    newWideCString("Highlight selected"),
    WS_CHILD or WS_VISIBLE or WS_TABSTOP,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idHighlight),
    hInst, nil)
  discard SendMessage(inspector.btnHighlight, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.btnExpand = CreateWindowExW(0, WC_BUTTON,
    newWideCString("Expand from selection"),
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

  inspector.uiaFilterLabel = CreateWindowExW(0, WC_STATIC,
    newWideCString("UIA filter (type, name, AutomationId):"),
    WS_CHILD or WS_VISIBLE, 0, 0, 0, 0, inspector.hwnd, HMENU(0), hInst, nil)
  discard SendMessage(inspector.uiaFilterLabel, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.uiaFilterEdit = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or WS_TABSTOP or ES_AUTOHSCROLL,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idUiaFilterEdit), hInst, nil)
  discard SendMessage(inspector.uiaFilterEdit, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

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

  resetWindowInfo(inspector)
  refreshWindowList(inspector)
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
  let depthText = if inspector.uiaMaxDepth > 0:
    fmt"Depth: {inspector.uiaMaxDepth}"
  else:
    "Depth: unknown"
  let pathText =
    if inspector.accPath.len > 0:
      fmt"{depthText} | Path: {inspector.accPath} (click to copy)"
    else:
      fmt"{depthText} | Acc path unavailable"
  discard SendMessage(inspector.statusBar, SB_SETTEXTW, 0, cast[LPARAM](newWideCString(versionText)))
  discard SendMessage(inspector.statusBar, SB_SETTEXTW, 1, cast[LPARAM](newWideCString(pathText)))

proc registerMenu(inspector: InspectorWindow) =
  let menuBar = CreateMenu()
  let viewMenu = CreateMenu()
  discard AppendMenuW(viewMenu, MF_STRING, idMenuHighlightColor,
    newWideCString("Highlight Color..."))
  discard AppendMenuW(menuBar, MF_POPUP, cast[UINT_PTR](viewMenu),
    newWideCString("&View"))
  discard SetMenu(inspector.hwnd, menuBar)

proc handleCommand(inspector: InspectorWindow; wParam: WPARAM; lParam: LPARAM) =
  let id = int(LOWORD(DWORD(wParam)))
  let code = int(HIWORD(DWORD(wParam)))
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
    let (tree, start) = inspector.expandTarget()
    beginExpandAll(inspector, tree, start)
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
    refreshWindowList(inspector)
  of idUiaFilterEdit:
    if code == EN_CHANGE:
      rebuildElementTree(inspector)
  else:
    discard

proc handleNotify(inspector: InspectorWindow; lParam: LPARAM) =
  let hdr = cast[ptr NMHDR](lParam)
  if hdr.hwndFrom == inspector.mainTree and hdr.code == UINT(TVN_SELCHANGEDW):
    let info = cast[ptr NMTREEVIEWW](lParam)
    let selectedElement = inspector.nodes.getOrDefault(info.itemNew.hItem, nil)
    if selectedElement != nil:
      discard highlightElementBounds(selectedElement, inspector.highlightColor, 1500,
        inspector.logger)
    populateProperties(inspector, selectedElement)
  elif hdr.hwndFrom == inspector.windowList and hdr.code == UINT(LVN_ITEMCHANGED):
    let info = cast[ptr NMLISTVIEW](lParam)
    if (info.uChanged and UINT(LVIF_STATE)) != 0 and
        (info.uNewState and UINT(LVIS_SELECTED)) != 0:
      handleWindowSelectionChanged(inspector)
  elif hdr.hwndFrom == inspector.propertiesList and hdr.code == UINT(NM_RCLICK):
    let info = cast[ptr NMITEMACTIVATE](lParam)
    if info.iItem >= 0 and info.iItem < inspector.propertyRows.len:
      let row = inspector.propertyRows[info.iItem]
      var text = if info.iSubItem == 0: row.name else: row.value
      if row.propertyId == UIA_ControlTypePropertyId:
        let sel = inspector.currentSelection()
        if not sel.isNil:
          text = controlTypeName(safeControlType(sel))
      if text.len > 0:
        copyToClipboard(text, inspector.logger)
  elif hdr.hwndFrom == inspector.patternsTree and hdr.code == UINT(NM_RCLICK):
    var pt: POINT
    discard GetCursorPos(addr pt)
    discard ScreenToClient(inspector.patternsTree, addr pt)
    var hit: TVHITTESTINFO
    hit.pt = pt
    let item = TreeView_HitTest(inspector.patternsTree, addr hit)
    if item != 0:
      discard TreeView_SelectItem(inspector.patternsTree, item)
      let text = inspector.patternCopyTexts.getOrDefault(item, "")
      if text.len > 0:
        copyToClipboard(text, inspector.logger)
  elif hdr.hwndFrom == inspector.patternsTree and hdr.code == UINT(NM_DBLCLK):
    let item = TreeView_GetSelection(inspector.patternsTree)
    if item != 0 and item in inspector.patternActions:
      let action = inspector.patternActions[item]
      executePatternAction(inspector, inspector.currentSelection(), action)
  elif hdr.hwndFrom == inspector.statusBar and (hdr.code == UINT(NM_CLICK) or hdr.code == UINT(NM_DBLCLK)):
    if inspector.accPath.len > 0:
      copyToClipboard(inspector.accPath, inspector.logger)

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
        handleCommand(inspector, wParam, lParam)
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
        syncFilterState(inspector)
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
