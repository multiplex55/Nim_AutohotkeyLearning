when system.hostOS != "windows":
  {.error: "UIA inspector window is only supported on Windows.".}

import std/[options, os, sequtils, strformat, strutils, tables]

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
  msgInitTree = WM_APP + 1
  contentPadding = 8
  groupPadding = 8
  buttonHeight = 26
  buttonSpacing = 6
  splitterWidth = 8
  minPanelWidth = 180
  minMiddleHeight = 100
  statusBarHeight = 24
  bottomPadding = 8
  contentBottomPadding = 12
  groupLabelPadding = 6
  expandTimerId = UINT_PTR(99)
  followMouseTimerId = UINT_PTR(101)
  followMouseIntervalMs = 280
  followMouseDebounceMs = 220

  idInvoke = 1001
  idSetFocus = 1002
  idHighlight = 1003
  idExpandAll = 1004
  idRefresh = 1005
  idUiaFilterEdit = 1006
  idRefreshTree = 1007
  idHighlightFollow = 1008
  idPropertiesList = 1100
  idPatternsTree = 1101
  idMainTree = 1102

  idStatusBar = 2002

  idGbWindowList = 3001
  idWindowFilterLabel = 3002
  idWindowFilterEdit = 3003
  idFilterVisibleCheck = 3004
  idFilterTitleCheck = 3005
  idFilterActivateCheck = 3006
  idWindowClassFilterLabel = 3007
  idWindowClassFilterEdit = 3008
  idWindowList = 3009

  idGbWindowInfo = 3020
  idWindowTitleLabel = 3021
  idWindowTitleValue = 3022
  idWindowHandleLabel = 3023
  idWindowHandleValue = 3024
  idWindowPosLabel = 3025
  idWindowPosValue = 3026
  idWindowClassInfoLabel = 3027
  idWindowClassValue = 3028
  idWindowProcessLabel = 3029
  idWindowProcessValue = 3030
  idWindowPidLabel = 3031
  idWindowPidValue = 3032
  idWindowAutomationIdLabel = 3033
  idWindowAutomationIdValue = 3034

  idGbProperties = 3040
  idGbPatterns = 3041
  idUiaFilterLabel = 3042

  idMenuHighlightColor = 2001

type
  PropertyRow = object
    name: string
    value: string
    propertyId: PROPERTYID

  PatternAction = object
    patternId: PATTERNID
    action: string

  ElementIdentifier = object
    path: seq[int]

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
    windowAutomationIdLabel: HWND
    windowAutomationIdValue: HWND
    filterVisible: HWND
    filterTitle: HWND
    filterActivate: HWND
    windowFilterLabel: HWND
    windowFilterEdit: HWND
    windowClassFilterLabel: HWND
    windowClassEdit: HWND
    btnRefresh: HWND
    btnTreeRefresh: HWND
    followHighlightCheck: HWND
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
    nodes: Table[HTREEITEM, ElementIdentifier]
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
    btnExpand: HWND
    uiaFilterLabel: HWND
    uiaFilterEdit: HWND
    uiaMaxDepth: int
    uiaFilterText: string
    propertyRows: seq[PropertyRow]
    patternActions: Table[HTREEITEM, PatternAction]
    patternCopyTexts: Table[HTREEITEM, string]
    accPath: string
    lastHoverHwnd: HWND
    lastHoverTick: DWORD
    lastActionContext: string

var inspectors = initTable[HWND, InspectorWindow]()
var commonControlsReady = false
proc updateStatusBar(inspector: InspectorWindow)
proc resetWindowInfo(inspector: InspectorWindow)
proc autoHighlight(inspector: InspectorWindow; element: ptr IUIAutomationElement)
proc nodeLabel(inspector: InspectorWindow; element: ptr IUIAutomationElement): string
proc currentSelectionId(inspector: InspectorWindow): Option[ElementIdentifier]
proc lParamX(lp: LPARAM): int =
  cast[int16](LOWORD(DWORD(lp))).int

proc lParamY(lp: LPARAM): int =
  cast[int16](HIWORD(DWORD(lp))).int

proc addContext(inspector: InspectorWindow; fields: var seq[(string, string)]) =
  if inspector != nil and inspector.lastActionContext.len > 0:
    fields.add(("context", inspector.lastActionContext))

proc formatPath(idOpt: Option[ElementIdentifier]): string =
  if idOpt.isNone:
    return "unknown"
  let parts = idOpt.get().path
  if parts.len == 0:
    "root"
  else:
    parts.mapIt($it).join("/")

proc updateActionContext(inspector: InspectorWindow; element: ptr IUIAutomationElement;
    idOpt: Option[ElementIdentifier] = none(ElementIdentifier)) =
  if inspector.isNil:
    return
  let label = if element.isNil: "Unavailable element" else: nodeLabel(inspector, element)
  inspector.lastActionContext = fmt"{label} (path: {formatPath(idOpt)})"

proc logComResult(inspector: InspectorWindow; message: string; hr: HRESULT) =
  if inspector.isNil or inspector.logger.isNil or hr == S_OK:
    return
  var fields = @[ ("hresult", fmt"0x{hr:X}") ]
  addContext(inspector, fields)
  inspector.logger.warn(message, fields)

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

proc safeLocalizedControlType(element: ptr IUIAutomationElement;
    inspector: InspectorWindow = nil): string =
  try:
    var val: VARIANT
    let hr = element.GetCurrentPropertyValue(UIA_LocalizedControlTypePropertyId,
      addr val)
    logComResult(inspector, "GetCurrentPropertyValue(LocalizedControlType) returned non-success", hr)
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

proc nodeLabel(inspector: InspectorWindow; element: ptr IUIAutomationElement): string =
  let name = safeCurrentName(element)
  let automationId = safeAutomationId(element)
  let localized = safeLocalizedControlType(element, inspector)
  let ctrlType = if localized.len > 0: localized else: controlTypeName(safeControlType(element))

  var parts: seq[string] = @[ctrlType]
  if name.len > 0:
    parts.add(&"\"{name}\"")
  if automationId.len > 0:
    parts.add(fmt"[{automationId}]")
  parts.join(" ")

proc releaseNodes(inspector: InspectorWindow) =
  inspector.nodes.clear()

proc lastErrorMessage(): string =
  let code = GetLastError()
  if code == 0:
    return "Unknown error"
  var buf: LPWSTR
  let flags = FORMAT_MESSAGE_ALLOCATE_BUFFER or FORMAT_MESSAGE_FROM_SYSTEM or
      FORMAT_MESSAGE_IGNORE_INSERTS
  let len = FormatMessageW(DWORD(flags), nil, code, 0, cast[LPWSTR](addr buf), 0, nil)
  if len == 0 or buf.isNil:
    return fmt"Win32 error {code}"
  defer: discard LocalFree(cast[HLOCAL](buf))
  let msg = $cast[WideCString](buf)
  fmt"Win32 error {code}: {msg.strip()}"

proc safeRootElement(inspector: InspectorWindow): ptr IUIAutomationElement =
  if inspector.uia.isNil:
    return nil
  try:
    inspector.uia.rootElement()
  except CatchableError as exc:
    if inspector.logger != nil:
      var fields = @[("error", exc.msg)]
      addContext(inspector, fields)
      inspector.logger.error("Failed to fetch UIA root element", fields)
    nil

proc elementFromId(inspector: InspectorWindow; id: ElementIdentifier): ptr IUIAutomationElement =
  try:
    if inspector.uia.isNil:
      return nil

    var walker: ptr IUIAutomationTreeWalker
    let hrWalker = inspector.uia.automation.get_ControlViewWalker(addr walker)
    var walkerHr = hrWalker
    if FAILED(hrWalker) or walker.isNil:
      walkerHr = inspector.uia.automation.get_RawViewWalker(addr walker)
    if FAILED(walkerHr) or walker.isNil:
      return nil
    defer: discard walker.Release()

    var current = inspector.safeRootElement()
    if current.isNil:
      return nil
    discard current.AddRef()

    for step in id.path:
      var child: ptr IUIAutomationElement
      let hrFirst = walker.GetFirstChildElement(current, addr child)
      logComResult(inspector, "GetFirstChildElement returned non-success", hrFirst)
      discard current.Release()
      if FAILED(hrFirst) or child.isNil:
        return nil

      var idx = 0
      var node = child
      while idx < step and node != nil:
        var next: ptr IUIAutomationElement
        let hrNext = walker.GetNextSiblingElement(node, addr next)
        logComResult(inspector, "GetNextSiblingElement returned non-success", hrNext)
        discard node.Release()
        if FAILED(hrNext) or hrNext == S_FALSE:
          return nil
        node = next
        inc idx

      if node.isNil:
        return nil
      current = node

    current
  except CatchableError:
    nil

proc controlTreeWalker(inspector: InspectorWindow): ptr IUIAutomationTreeWalker

proc elementsEqual(inspector: InspectorWindow; a, b: ptr IUIAutomationElement): bool =
  if inspector.isNil or inspector.uia.isNil or a.isNil or b.isNil:
    return false
  var same: BOOL = 0
  let hr = inspector.uia.automation.CompareElements(a, b, addr same)
  logComResult(inspector, "CompareElements returned non-success", hr)
  SUCCEEDED(hr) and same != 0

proc findElementPath(inspector: InspectorWindow; target: ptr IUIAutomationElement;
    walker: ptr IUIAutomationTreeWalker; current: ptr IUIAutomationElement;
    path: seq[int]): Option[ElementIdentifier] =
  if target.isNil or walker.isNil or current.isNil:
    return
  if elementsEqual(inspector, current, target):
    return some(ElementIdentifier(path: path))

  var child: ptr IUIAutomationElement
  let hrFirst = walker.GetFirstChildElement(current, addr child)
  logComResult(inspector, "GetFirstChildElement returned non-success", hrFirst)
  if FAILED(hrFirst) or child.isNil:
    return

  var idx = 0
  var node = child
  while node != nil:
    let found = findElementPath(inspector, target, walker, node, path & @[idx])

    var next: ptr IUIAutomationElement
    let hrNext = walker.GetNextSiblingElement(node, addr next)
    logComResult(inspector, "GetNextSiblingElement returned non-success", hrNext)
    discard node.Release()

    if found.isSome:
      if not next.isNil:
        discard next.Release()
      return found

    if FAILED(hrNext) or hrNext == S_FALSE:
      break
    node = next
    inc idx

proc elementIdentifier(inspector: InspectorWindow; element: ptr IUIAutomationElement): Option[ElementIdentifier] =
  if inspector.uia.isNil or element.isNil:
    return

  let walker = controlTreeWalker(inspector)
  if walker.isNil:
    return
  defer: discard walker.Release()

  var root = inspector.safeRootElement()
  if root.isNil:
    return
  discard root.AddRef()
  defer: discard root.Release()

  findElementPath(inspector, element, walker, root, @[])

proc rebuildElementTree(inspector: InspectorWindow)
proc refreshTreeItemChildren(inspector: InspectorWindow; item: HTREEITEM)
proc populateProperties(inspector: InspectorWindow; element: ptr IUIAutomationElement)
proc followHighlightEnabled(inspector: InspectorWindow): bool

proc ensureTreePath(inspector: InspectorWindow; id: ElementIdentifier): Option[HTREEITEM] =
  if inspector.mainTree == 0:
    return

  var rootItem = TreeView_GetRoot(inspector.mainTree)
  if rootItem == 0:
    rebuildElementTree(inspector)
    rootItem = TreeView_GetRoot(inspector.mainTree)
  if rootItem == 0:
    return

  if id.path.len == 0:
    return some(rootItem)

  var current = rootItem
  for childIndex in id.path:
    refreshTreeItemChildren(inspector, current)
    discard TreeView_Expand(inspector.mainTree, current, UINT(TVE_EXPAND))

    var child = TreeView_GetChild(inspector.mainTree, current)
    var idx = 0
    while child != 0 and idx < childIndex:
      child = TreeView_GetNextSibling(inspector.mainTree, child)
      inc idx

    if child == 0:
      return
    current = child

  some(current)

proc screenPointFromMsg(inspector: InspectorWindow; hwnd: HWND; msg: UINT;
    lParam: LPARAM): Option[POINT] =
  var pt: POINT
  case msg
  of WM_CONTEXTMENU:
    if lParam == LPARAM(-1):
      discard GetCursorPos(addr pt)
    else:
      pt.x = LONG(lParamX(lParam))
      pt.y = LONG(lParamY(lParam))
    some(pt)
  of WM_RBUTTONUP:
    pt.x = LONG(lParamX(lParam))
    pt.y = LONG(lParamY(lParam))
    discard ClientToScreen(hwnd, addr pt)
    some(pt)
  else:
    none(POINT)

proc selectElementUnderCursor(inspector: InspectorWindow; screenPt: POINT) =
  if inspector.isNil or inspector.uia.isNil or not inspector.followHighlightEnabled():
    return

  let targetHwnd = WindowFromPoint(screenPt)
  if targetHwnd == 0 or targetHwnd == inspector.hwnd or IsChild(inspector.hwnd, targetHwnd) != 0:
    return

  try:
    let element = inspector.uia.fromPoint(screenPt.x, screenPt.y)
    if element.isNil:
      return
    defer: discard element.Release()

    let idOpt = inspector.elementIdentifier(element)
    updateActionContext(inspector, element, idOpt)
    if idOpt.isNone:
      return

    let treeItem = inspector.ensureTreePath(idOpt.get())
    if treeItem.isNone or treeItem.get() == 0:
      return

    discard TreeView_SelectItem(inspector.mainTree, treeItem.get())
    populateProperties(inspector, element)
  except CatchableError as exc:
    if inspector.logger != nil:
      var fields = @[("error", exc.msg)]
      addContext(inspector, fields)
      inspector.logger.debug("Follow-highlight selection failed", fields)

proc setTreeItemBold(tree: HWND; item: HTREEITEM; bold: bool) =
  var tvi: TVITEMW
  tvi.mask = UINT(TVIF_STATE)
  tvi.hItem = item
  tvi.stateMask = UINT(TVIS_BOLD)
  tvi.state = if bold: UINT(TVIS_BOLD) else: 0
  discard TreeView_SetItem(tree, addr tvi)

proc addTreeItem(tree: HWND; parent: HTREEITEM; text: string;
    data: LPARAM = 0; hasChildren: bool = false): HTREEITEM =
  var insert: TVINSERTSTRUCTW
  insert.hParent = parent
  insert.hInsertAfter = TVI_LAST
  insert.item.mask = UINT(TVIF_TEXT or TVIF_PARAM)
  if hasChildren:
    insert.item.mask = insert.item.mask or UINT(TVIF_CHILDREN)
    insert.item.cChildren = 1
  let wide = newWideCString(text)
  insert.item.pszText = wide
  insert.item.cchTextMax = int32(text.len)
  insert.item.lParam = data
  TreeView_InsertItem(tree, addr insert)

proc elementMatchesFilter(inspector: InspectorWindow; element: ptr IUIAutomationElement;
    filterLower: string): bool =
  if element.isNil or filterLower.len == 0:
    return false

  let localized = safeLocalizedControlType(element, inspector).toLower()
  let ctrlTypeFallback = controlTypeName(safeControlType(element)).toLower()
  let name = safeCurrentName(element).toLower()
  let automationId = safeAutomationId(element).toLower()

  for candidate in [localized, ctrlTypeFallback, name, automationId]:
    if candidate.len > 0 and candidate.find(filterLower) >= 0:
      return true
  false

proc controlTreeWalker(inspector: InspectorWindow): ptr IUIAutomationTreeWalker =
  if inspector.uia.isNil:
    return nil
  var walker: ptr IUIAutomationTreeWalker
  let hrWalker = inspector.uia.automation.get_ControlViewWalker(addr walker)
  logComResult(inspector, "get_ControlViewWalker returned non-success", hrWalker)
  if FAILED(hrWalker) or walker.isNil:
    return nil
  walker

proc elementHasChildren(inspector: InspectorWindow; walker: ptr IUIAutomationTreeWalker;
    element: ptr IUIAutomationElement): bool =
  if walker.isNil or element.isNil:
    return false
  var child: ptr IUIAutomationElement
  let hr = walker.GetFirstChildElement(element, addr child)
  logComResult(inspector, "GetFirstChildElement returned non-success", hr)
  if not child.isNil:
    discard child.Release()
  not child.isNil

proc removeNodeMappings(inspector: InspectorWindow; item: HTREEITEM) =
  var child = TreeView_GetChild(inspector.mainTree, item)
  while child != 0:
    let next = TreeView_GetNextSibling(inspector.mainTree, child)
    removeNodeMappings(inspector, child)
    inspector.nodes.del(child)
    child = next

proc clearNodeChildren(inspector: InspectorWindow; parent: HTREEITEM) =
  var child = TreeView_GetChild(inspector.mainTree, parent)
  while child != 0:
    let next = TreeView_GetNextSibling(inspector.mainTree, child)
    removeNodeMappings(inspector, child)
    inspector.nodes.del(child)
    discard TreeView_DeleteItem(inspector.mainTree, child)
    child = next

proc populateChildNodes(inspector: InspectorWindow; parentItem: HTREEITEM;
    parentId: ElementIdentifier; walker: ptr IUIAutomationTreeWalker;
    filterLower: string; hasFilter: bool) =
  clearNodeChildren(inspector, parentItem)
  if walker.isNil:
    return

  var parentElement: ptr IUIAutomationElement
  if parentId.path.len == 0:
    parentElement = inspector.safeRootElement()
    if parentElement.isNil:
      return
    discard parentElement.AddRef()
  else:
    parentElement = inspector.elementFromId(parentId)
    if parentElement.isNil:
      return
  defer: discard parentElement.Release()

  var child: ptr IUIAutomationElement
  let hrFirst = walker.GetFirstChildElement(parentElement, addr child)
  logComResult(inspector, "GetFirstChildElement returned non-success", hrFirst)
  if FAILED(hrFirst) or child.isNil:
    return

  var current = child
  var childIndex = 0
  while current != nil:
    let childHasChildren = elementHasChildren(inspector, walker, current)
    let label = nodeLabel(inspector, current)
    let id = ElementIdentifier(path: parentId.path & @[childIndex])
    let item = addTreeItem(inspector.mainTree, parentItem, label,
      hasChildren = childHasChildren)
    inspector.nodes[item] = id
    if hasFilter:
      setTreeItemBold(inspector.mainTree, item, elementMatchesFilter(inspector, current,
        filterLower))
    inspector.uiaMaxDepth = max(inspector.uiaMaxDepth, id.path.len + 1)

    var next: ptr IUIAutomationElement
    let hrNext = walker.GetNextSiblingElement(current, addr next)
    logComResult(inspector, "GetNextSiblingElement returned non-success", hrNext)
    discard current.Release()
    inc childIndex
    if FAILED(hrNext) or hrNext == S_FALSE:
      break
    current = next

proc rebuildElementTree(inspector: InspectorWindow) =
  TreeView_DeleteAllItems(inspector.mainTree)
  releaseNodes(inspector)
  inspector.uiaMaxDepth = 0

  if inspector.uia.isNil:
    if inspector.logger != nil:
      inspector.logger.error("UIA inspector missing automation instance")
    return
  let walker = controlTreeWalker(inspector)
  if walker.isNil:
    if inspector.logger != nil:
      inspector.logger.error("Failed to create UIA control view walker")
    return
  defer: discard walker.Release()

  var root = inspector.safeRootElement()
  if root.isNil:
    if inspector.logger != nil:
      inspector.logger.warn("UIA root element unavailable; cannot build inspector tree")
    return

  inspector.uiaFilterText =
    if inspector.uiaFilterEdit != 0: readEditText(inspector.uiaFilterEdit).strip()
    else: ""
  let filterLower = inspector.uiaFilterText.toLower()
  let hasFilter = filterLower.len > 0

  let rootHasChildren = elementHasChildren(inspector, walker, root)
  let rootItem = addTreeItem(inspector.mainTree, TVI_ROOT, nodeLabel(inspector, root),
    hasChildren = rootHasChildren)
  inspector.nodes[rootItem] = ElementIdentifier(path: @[])
  if hasFilter:
    setTreeItemBold(inspector.mainTree, rootItem, elementMatchesFilter(inspector, root,
      filterLower))
  inspector.uiaMaxDepth = 1

  discard TreeView_SelectItem(inspector.mainTree, rootItem)
  if rootHasChildren:
    TreeView_Expand(inspector.mainTree, rootItem, UINT(TVE_EXPAND))
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
    updateActionContext(inspector, element, some(ElementIdentifier(path: @[])))
    inspector.uia.setRootElement(element)
    rebuildElementTree(inspector)
    autoHighlight(inspector, element)
  except CatchableError as exc:
    if inspector.logger != nil:
      var fields = @[("error", exc.msg)]
      addContext(inspector, fields)
      inspector.logger.error("Failed to build inspector tree from window", fields)
    inspector.uia.setRootElement(nil)
  resetWindowInfo(inspector)

proc setEditText(hwnd: HWND; value: string) =
  discard SetWindowTextW(hwnd, newWideCString(value))

proc resetWindowInfo(inspector: InspectorWindow) =
  setEditText(inspector.windowTitleValue, "No selection")
  for hwnd in [inspector.windowHandleValue, inspector.windowPosValue,
      inspector.windowAutomationIdValue, inspector.windowClassValue,
      inspector.windowProcessValue, inspector.windowPidValue]:
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

proc propertyValueString(inspector: InspectorWindow; element: ptr IUIAutomationElement;
    propertyId: PROPERTYID): string =
  if element.isNil:
    return ""
  if propertyId == UIA_ControlTypePropertyId:
    return controlTypeName(safeControlType(element))
  if propertyId == UIA_LocalizedControlTypePropertyId:
    return safeLocalizedControlType(element, inspector)
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
  logComResult(inspector, "GetCurrentPropertyValue returned non-success", hr)
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
    let value = propertyValueString(inspector, element, prop[1])
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
  setEditText(inspector.windowAutomationIdValue,
    if automationId.len > 0: automationId else: "Unavailable")
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
  logComResult(inspector, "GetCurrentPropertyValue(ProcessId) returned non-success", hr)
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

proc getPattern[T](inspector: InspectorWindow; element: ptr IUIAutomationElement;
    patternId: PATTERNID; resultPtr: var ptr T): bool =
  resultPtr = nil
  if element.isNil:
    return false
  let hr = element.GetCurrentPattern(patternId,
    cast[ptr ptr IUnknown](addr resultPtr))
  logComResult(inspector, "GetCurrentPattern returned non-success", hr)
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
  if getPattern(inspector, element, UIA_InvokePatternId, invoke):
    defer: discard invoke.Release()
    let root = addPatternNode(inspector, TVI_ROOT, "Invoke", "Invoke")
    discard addPatternNode(inspector, root, "Action: Invoke", "Invoke",
      PatternAction(patternId: UIA_InvokePatternId, action: "Invoke"))
    discard TreeView_Expand(inspector.patternsTree, root, UINT(TVE_EXPAND))
    anyPattern = true

  proc bstrValue(text: BSTR; hr: HRESULT): string =
    if text.isNil or FAILED(hr):
      ""
    else:
      $cast[WideCString](text)

  var legacy: ptr IUIAutomationLegacyIAccessiblePattern
  if getPattern(inspector, element, UIA_LegacyIAccessiblePatternId, legacy):
    defer: discard legacy.Release()
    var name: BSTR = nil
    var value: BSTR = nil
    var desc: BSTR = nil
    var role: LONG = 0
    let hrName = legacy.get_CurrentName(addr name)
    let hrValue = legacy.get_CurrentValue(addr value)
    let hrDesc = legacy.get_CurrentDescription(addr desc)
    let hrRole = legacy.get_CurrentRole(addr role)
    logComResult(inspector, "LegacyIAccessible get_CurrentName returned non-success", hrName)
    logComResult(inspector, "LegacyIAccessible get_CurrentValue returned non-success", hrValue)
    logComResult(inspector, "LegacyIAccessible get_CurrentDescription returned non-success", hrDesc)
    logComResult(inspector, "LegacyIAccessible get_CurrentRole returned non-success", hrRole)
    let root = addPatternNode(inspector, TVI_ROOT, "LegacyIAccessible", "LegacyIAccessible")
    discard addPatternNode(inspector, root, "CurrentName: " &
      bstrValue(name, hrName),
      bstrValue(name, hrName))
    discard addPatternNode(inspector, root, "CurrentValue: " &
      bstrValue(value, hrValue),
      bstrValue(value, hrValue))
    discard addPatternNode(inspector, root, "CurrentDescription: " &
      bstrValue(desc, hrDesc),
      bstrValue(desc, hrDesc))
    discard addPatternNode(inspector, root, "CurrentRole: " & (if FAILED(hrRole): "Unavailable" else: roleText(role)),
      if FAILED(hrRole): "Unavailable" else: roleText(role))
    discard addPatternNode(inspector, root, "Action: DoDefaultAction", "DoDefaultAction",
      PatternAction(patternId: UIA_LegacyIAccessiblePatternId, action: "DoDefaultAction"))
    if not name.isNil and SUCCEEDED(hrName): SysFreeString(name)
    if not value.isNil and SUCCEEDED(hrValue): SysFreeString(value)
    if not desc.isNil and SUCCEEDED(hrDesc): SysFreeString(desc)
    discard TreeView_Expand(inspector.patternsTree, root, UINT(TVE_EXPAND))
    anyPattern = true

  var selectionItem: ptr IUIAutomationSelectionItemPattern
  if getPattern(inspector, element, UIA_SelectionItemPatternId, selectionItem):
    defer: discard selectionItem.Release()
    var isSelected: BOOL = 0
    let hrSelected = selectionItem.get_CurrentIsSelected(addr isSelected)
    logComResult(inspector, "SelectionItem get_CurrentIsSelected returned non-success", hrSelected)
    let root = addPatternNode(inspector, TVI_ROOT, "SelectionItem", "SelectionItem")
    addBoolChild(inspector, root, "CurrentIsSelected", isSelected != 0)
    discard addPatternNode(inspector, root, "Action: Select", "Select",
      PatternAction(patternId: UIA_SelectionItemPatternId, action: "Select"))
    discard TreeView_Expand(inspector.patternsTree, root, UINT(TVE_EXPAND))
    anyPattern = true

  var valuePattern: ptr IUIAutomationValuePattern
  if getPattern(inspector, element, UIA_ValuePatternId, valuePattern):
    defer: discard valuePattern.Release()
    var current: BSTR = nil
    var readOnly: BOOL = 0
    let hrCurrent = valuePattern.get_CurrentValue(addr current)
    let hrReadOnly = valuePattern.get_CurrentIsReadOnly(addr readOnly)
    logComResult(inspector, "ValuePattern get_CurrentValue returned non-success", hrCurrent)
    logComResult(inspector, "ValuePattern get_CurrentIsReadOnly returned non-success", hrReadOnly)
    let root = addPatternNode(inspector, TVI_ROOT, "Value", "Value")
    discard addPatternNode(inspector, root, "CurrentValue: " &
      bstrValue(current, hrCurrent),
      bstrValue(current, hrCurrent))
    addBoolChild(inspector, root, "CurrentIsReadOnly",
      hrReadOnly == S_OK and readOnly != 0)
    discard addPatternNode(inspector, root, "Action: SetValue (uses clipboard text)", "SetValue",
      PatternAction(patternId: UIA_ValuePatternId, action: "SetValue"))
    if not current.isNil and SUCCEEDED(hrCurrent): SysFreeString(current)
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
  updateActionContext(inspector, element, inspector.currentSelectionId())
  case action.action
  of "Invoke":
    var pattern: ptr IUIAutomationInvokePattern
    if getPattern(inspector, element, UIA_InvokePatternId, pattern):
      defer: discard pattern.Release()
      let hr = pattern.Invoke()
      logComResult(inspector, "Invoke pattern call returned non-success", hr)
  of "DoDefaultAction":
    var pattern: ptr IUIAutomationLegacyIAccessiblePattern
    if getPattern(inspector, element, UIA_LegacyIAccessiblePatternId, pattern):
      defer: discard pattern.Release()
      let hr = pattern.DoDefaultAction()
      logComResult(inspector, "DoDefaultAction returned non-success", hr)
  of "Select":
    var pattern: ptr IUIAutomationSelectionItemPattern
    if getPattern(inspector, element, UIA_SelectionItemPatternId, pattern):
      defer: discard pattern.Release()
      let hr = pattern.Select()
      logComResult(inspector, "SelectionItem.Select returned non-success", hr)
  of "SetValue":
    var pattern: ptr IUIAutomationValuePattern
    if getPattern(inspector, element, UIA_ValuePatternId, pattern):
      defer: discard pattern.Release()
      let clip = readClipboardText()
      if clip.isSome:
        let text = clip.get()
        let wide = newWideCString(text)
        let bstr = SysAllocStringLen(cast[ptr WCHAR](addr wide[0]), UINT(text.len))
        if bstr != nil:
          let hr = pattern.SetValue(bstr)
          logComResult(inspector, "Value.SetValue returned non-success", hr)
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
  let selectionId = inspector.currentSelectionId()
  if not element.isNil:
    updateActionContext(inspector, element,
      if selectionId.isSome: selectionId else: some(ElementIdentifier(path: @[])))
  else:
    updateActionContext(inspector, element, selectionId)
  if element.isNil:
    discard EnableWindow(inspector.btnInvoke, FALSE)
    discard EnableWindow(inspector.btnFocus, FALSE)
    discard EnableWindow(inspector.btnHighlight, FALSE)
    discard EnableWindow(inspector.btnExpand, FALSE)
    populatePropertyList(inspector, nil)
    populatePatterns(inspector, nil)
    resetWindowInfo(inspector)
    updateStatusBar(inspector)
    return

  discard EnableWindow(inspector.btnInvoke, TRUE)
  discard EnableWindow(inspector.btnFocus, TRUE)
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

proc currentSelectionId(inspector: InspectorWindow): Option[ElementIdentifier] =
  let selected = TreeView_GetSelection(inspector.mainTree)
  if selected == 0 or selected notin inspector.nodes:
    return
  some(inspector.nodes[selected])

proc elementFromSelection(inspector: InspectorWindow): ptr IUIAutomationElement =
  let idOpt = inspector.currentSelectionId()
  if idOpt.isNone:
    return nil
  inspector.elementFromId(idOpt.get())

proc refreshTreeItemChildren(inspector: InspectorWindow; item: HTREEITEM) =
  if inspector.isNil or item == 0 or inspector.mainTree == 0:
    return
  if item notin inspector.nodes:
    return

  let walker = controlTreeWalker(inspector)
  if walker.isNil:
    if inspector.logger != nil:
      inspector.logger.error("Failed to create control view walker for expansion")
    return
  defer: discard walker.Release()

  let filterLower = inspector.uiaFilterText.toLower()
  populateChildNodes(inspector, item, inspector.nodes[item], walker, filterLower,
    filterLower.len > 0)
  updateStatusBar(inspector)

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
  let element = inspector.elementFromSelection()
  if element.isNil:
    return
  defer: discard element.Release()
  updateActionContext(inspector, element, inspector.currentSelectionId())
  try:
    inspector.uia.invoke(element)
    if inspector.logger != nil:
      inspector.logger.info("Invoked UIA element",
        [("name", safeCurrentName(element)), ("automationId", safeAutomationId(element))])
  except CatchableError as exc:
    if inspector.logger != nil:
      var fields = @[("error", exc.msg)]
      addContext(inspector, fields)
      inspector.logger.error("UIA invoke failed", fields)

proc handleSetFocus(inspector: InspectorWindow) =
  let element = inspector.elementFromSelection()
  if element.isNil:
    return
  defer: discard element.Release()
  updateActionContext(inspector, element, inspector.currentSelectionId())
  try:
    let hr = element.SetFocus()
    if FAILED(hr):
      raise newException(UiaError, fmt"SetFocus failed (0x{hr:X})")
    if inspector.logger != nil:
      inspector.logger.info("Set keyboard focus to element",
        [("name", safeCurrentName(element)), ("automationId", safeAutomationId(element))])
  except CatchableError as exc:
    if inspector.logger != nil:
      var fields = @[("error", exc.msg)]
      addContext(inspector, fields)
      inspector.logger.error("Failed to set focus", fields)

proc handleHighlight(inspector: InspectorWindow) =
  let element = inspector.elementFromSelection()
  if element.isNil:
    discard EnableWindow(inspector.btnHighlight, FALSE)
    return
  defer: discard element.Release()
  updateActionContext(inspector, element, inspector.currentSelectionId())

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
  let contentHeight = max(0, usableHeight - contentBottomPadding)
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
    bottom: LONG(contentTop + contentHeight)
  )
  inspector.splitters[1] = RECT(
    left: LONG(middleX + middleWidth),
    right: LONG(middleX + middleWidth + splitterWidth),
    top: LONG(contentTop),
    bottom: LONG(contentTop + contentHeight)
  )

  MoveWindow(inspector.gbWindowList, leftX.cint, contentTop.cint,
    leftWidth.int32, contentHeight.int32, TRUE)

  let groupInnerLeft = leftX + groupPadding
  var currentY = contentTop + groupPadding + groupLabelPadding
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
    max(contentHeight - (currentY - contentTop) - groupPadding, 80).int32, TRUE)

  let middleHeight = contentHeight
  var infoHeight = inspector.state.infoHeight
  let maxInfo = max(minSection, middleHeight - splitterWidth - minSection)
  let minInfo = minSection
  infoHeight = clamp(infoHeight, minInfo, maxInfo)
  inspector.state.infoHeight = infoHeight
  MoveWindow(inspector.gbWindowInfo, middleX.cint, contentTop.cint,
    middleWidth.int32, infoHeight.int32, TRUE)

  let infoInnerWidth = middleWidth - 2 * groupPadding
  var infoY = contentTop + groupPadding + groupLabelPadding
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

  MoveWindow(inspector.windowAutomationIdLabel, (middleX + groupPadding).cint, infoY.cint,
    labelWidth.int32, rowHeight.int32, TRUE)
  MoveWindow(inspector.windowAutomationIdValue,
    (middleX + groupPadding + labelWidth + 4).cint,
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

  let propInnerY = propertiesY + groupPadding + groupLabelPadding
  let propInnerHeight = propBoxHeight - 2 * groupPadding - groupLabelPadding
  MoveWindow(inspector.propertiesList, (middleX + groupPadding).cint,
    propInnerY.cint, (middleWidth - 2 * groupPadding).int32,
    max(propInnerHeight, 20).int32, TRUE)

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
    (patternsY + groupPadding + groupLabelPadding).cint, (middleWidth - 2 * groupPadding).int32,
    patternsInnerHeight.int32, TRUE)

  let patternBtnY = max(patternsY + patternBoxHeight - groupPadding - buttonHeight,
    patternsY + groupPadding + groupLabelPadding)
  var patternBtnX = middleX + groupPadding
  MoveWindow(inspector.btnInvoke, patternBtnX.cint, patternBtnY.cint, 100, buttonHeight.int32, TRUE)
  patternBtnX += 100 + buttonSpacing
  MoveWindow(inspector.btnFocus, patternBtnX.cint, patternBtnY.cint, 100, buttonHeight.int32, TRUE)

  let filterLabelHeight = 16
  let filterEditHeight = 22
  var rightControlsY = contentTop + groupPadding
  var rightControlX = rightX + groupPadding
  MoveWindow(inspector.btnTreeRefresh, rightControlX.cint, rightControlsY.cint,
    120, buttonHeight.int32, TRUE)
  rightControlX += 120 + buttonSpacing
  MoveWindow(inspector.btnHighlight, rightControlX.cint, rightControlsY.cint, 140, buttonHeight.int32,
    TRUE)
  rightControlX += 140 + buttonSpacing
  MoveWindow(inspector.btnExpand, rightControlX.cint, rightControlsY.cint, 160, buttonHeight.int32,
    TRUE)
  rightControlsY += buttonHeight + groupPadding
  rightControlX = rightX + groupPadding
  MoveWindow(inspector.followHighlightCheck, rightControlX.cint, rightControlsY.cint,
    max(0, rightWidth - 2 * groupPadding).int32, buttonHeight.int32, TRUE)
  rightControlsY += buttonHeight + groupPadding

  let filterTop = rightControlsY
  MoveWindow(inspector.uiaFilterLabel, (rightX + groupPadding).cint, filterTop.cint,
    (rightWidth - 2 * groupPadding).int32, filterLabelHeight.int32, TRUE)
  MoveWindow(inspector.uiaFilterEdit, (rightX + groupPadding).cint,
    (filterTop + filterLabelHeight + 4).cint,
    (rightWidth - 2 * groupPadding).int32, filterEditHeight.int32, TRUE)

  let treeTop = filterTop + filterLabelHeight + filterEditHeight + 8
  let treeHeight = max(contentHeight - (treeTop - contentTop), 40)
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
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idGbWindowList), hInst, nil)
  discard SendMessage(inspector.gbWindowList, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowFilterLabel = CreateWindowExW(0, WC_STATIC,
    newWideCString("Title filter:"), WS_CHILD or WS_VISIBLE,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idWindowFilterLabel), hInst, nil)
  discard SendMessage(inspector.windowFilterLabel, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowFilterEdit = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or WS_TABSTOP or ES_AUTOHSCROLL,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idWindowFilterEdit), hInst, nil)
  discard SendMessage(inspector.windowFilterEdit, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.filterVisible = CreateWindowExW(0, WC_BUTTON,
    newWideCString("Visible"), WS_CHILD or WS_VISIBLE or WS_TABSTOP or BS_AUTOCHECKBOX,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idFilterVisibleCheck), hInst, nil)
  discard SendMessage(inspector.filterVisible, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  discard SendMessage(inspector.filterVisible, BM_SETCHECK,
    WPARAM(if inspector.state.filterVisible: BST_CHECKED else: BST_UNCHECKED), 0)

  inspector.filterTitle = CreateWindowExW(0, WC_BUTTON,
    newWideCString("Title"), WS_CHILD or WS_VISIBLE or WS_TABSTOP or BS_AUTOCHECKBOX,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idFilterTitleCheck), hInst, nil)
  discard SendMessage(inspector.filterTitle, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  discard SendMessage(inspector.filterTitle, BM_SETCHECK,
    WPARAM(if inspector.state.filterTitle: BST_CHECKED else: BST_UNCHECKED), 0)

  inspector.filterActivate = CreateWindowExW(0, WC_BUTTON,
    newWideCString("Activate"), WS_CHILD or WS_VISIBLE or WS_TABSTOP or BS_AUTOCHECKBOX,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idFilterActivateCheck), hInst, nil)
  discard SendMessage(inspector.filterActivate, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  discard SendMessage(inspector.filterActivate, BM_SETCHECK,
    WPARAM(if inspector.state.filterActivate: BST_CHECKED else: BST_UNCHECKED), 0)

  inspector.windowClassFilterLabel = CreateWindowExW(0, WC_STATIC,
    newWideCString("Class filter:"), WS_CHILD or WS_VISIBLE,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idWindowClassFilterLabel), hInst, nil)
  discard SendMessage(inspector.windowClassFilterLabel, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowClassEdit = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or WS_TABSTOP or ES_AUTOHSCROLL,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idWindowClassFilterEdit), hInst, nil)
  discard SendMessage(inspector.windowClassEdit, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.btnRefresh = CreateWindowExW(0, WC_BUTTON, newWideCString("Refresh window list"),
    WS_CHILD or WS_VISIBLE or WS_TABSTOP,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idRefresh), hInst, nil)
  discard SendMessage(inspector.btnRefresh, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.btnTreeRefresh = CreateWindowExW(0, WC_BUTTON, newWideCString("Refresh tree"),
    WS_CHILD or WS_VISIBLE or WS_TABSTOP,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idRefreshTree), hInst, nil)
  discard SendMessage(inspector.btnTreeRefresh, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.followHighlightCheck = CreateWindowExW(0, WC_BUTTON,
    newWideCString("Highlight follow mouse"),
    WS_CHILD or WS_VISIBLE or WS_TABSTOP or BS_AUTOCHECKBOX,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idHighlightFollow), hInst, nil)
  discard SendMessage(inspector.followHighlightCheck, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  discard SendMessage(inspector.followHighlightCheck, BM_SETCHECK,
    WPARAM(if inspector.state.highlightFollow: BST_CHECKED else: BST_UNCHECKED), 0)

  var listStyle = WS_CHILD or WS_VISIBLE or WS_TABSTOP or LVS_REPORT or LVS_SHOWSELALWAYS or
      LVS_SINGLESEL or WS_BORDER
  inspector.windowList = CreateWindowExW(DWORD(WS_EX_CLIENTEDGE), WC_LISTVIEWW, nil,
    DWORD(listStyle), 0, 0, 0, 0, inspector.hwnd, cast[HMENU](idWindowList), hInst, nil)
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
    WS_CHILD or WS_VISIBLE or BS_GROUPBOX, 0, 0, 0, 0, inspector.hwnd, cast[HMENU](idGbWindowInfo), hInst, nil)
  discard SendMessage(inspector.gbWindowInfo, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowTitleLabel = CreateWindowExW(0, WC_STATIC, newWideCString("Title:"),
    WS_CHILD or WS_VISIBLE, 0, 0, 0, 0, inspector.hwnd, cast[HMENU](idWindowTitleLabel), hInst, nil)
  discard SendMessage(inspector.windowTitleLabel, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  inspector.windowTitleValue = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or ES_AUTOHSCROLL or ES_READONLY,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idWindowTitleValue), hInst, nil)
  discard SendMessage(inspector.windowTitleValue, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowHandleLabel = CreateWindowExW(0, WC_STATIC, newWideCString("HWND:"),
    WS_CHILD or WS_VISIBLE, 0, 0, 0, 0, inspector.hwnd, cast[HMENU](idWindowHandleLabel), hInst, nil)
  discard SendMessage(inspector.windowHandleLabel, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  inspector.windowHandleValue = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or ES_AUTOHSCROLL or ES_READONLY,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idWindowHandleValue), hInst, nil)
  discard SendMessage(inspector.windowHandleValue, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowPosLabel = CreateWindowExW(0, WC_STATIC, newWideCString("Position:"),
    WS_CHILD or WS_VISIBLE, 0, 0, 0, 0, inspector.hwnd, cast[HMENU](idWindowPosLabel), hInst, nil)
  discard SendMessage(inspector.windowPosLabel, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  inspector.windowPosValue = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or ES_AUTOHSCROLL or ES_READONLY,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idWindowPosValue), hInst, nil)
  discard SendMessage(inspector.windowPosValue, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowClassInfoLabel = CreateWindowExW(0, WC_STATIC, newWideCString("Class:"),
    WS_CHILD or WS_VISIBLE, 0, 0, 0, 0, inspector.hwnd, cast[HMENU](idWindowClassInfoLabel), hInst, nil)
  discard SendMessage(inspector.windowClassInfoLabel, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  inspector.windowClassValue = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or ES_AUTOHSCROLL or ES_READONLY,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idWindowClassValue), hInst, nil)
  discard SendMessage(inspector.windowClassValue, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowProcessLabel = CreateWindowExW(0, WC_STATIC, newWideCString("Process:"),
    WS_CHILD or WS_VISIBLE, 0, 0, 0, 0, inspector.hwnd, cast[HMENU](idWindowProcessLabel), hInst, nil)
  discard SendMessage(inspector.windowProcessLabel, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  inspector.windowProcessValue = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or ES_AUTOHSCROLL or ES_READONLY,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idWindowProcessValue), hInst, nil)
  discard SendMessage(inspector.windowProcessValue, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowPidLabel = CreateWindowExW(0, WC_STATIC, newWideCString("PID:"),
    WS_CHILD or WS_VISIBLE, 0, 0, 0, 0, inspector.hwnd, cast[HMENU](idWindowPidLabel), hInst, nil)
  discard SendMessage(inspector.windowPidLabel, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  inspector.windowPidValue = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or ES_AUTOHSCROLL or ES_READONLY,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idWindowPidValue), hInst, nil)
  discard SendMessage(inspector.windowPidValue, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.windowAutomationIdLabel = CreateWindowExW(0, WC_STATIC,
    newWideCString("AutomationId:"), WS_CHILD or WS_VISIBLE,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idWindowAutomationIdLabel), hInst, nil)
  discard SendMessage(inspector.windowAutomationIdLabel, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  inspector.windowAutomationIdValue = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or ES_AUTOHSCROLL or ES_READONLY,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idWindowAutomationIdValue), hInst, nil)
  discard SendMessage(inspector.windowAutomationIdValue, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.gbProperties = CreateWindowExW(0, WC_BUTTON, newWideCString("Properties"),
    WS_CHILD or WS_VISIBLE or BS_GROUPBOX, 0, 0, 0, 0, inspector.hwnd, cast[HMENU](idGbProperties), hInst, nil)
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
    WS_CHILD or WS_VISIBLE or BS_GROUPBOX, 0, 0, 0, 0, inspector.hwnd, cast[HMENU](idGbPatterns), hInst, nil)
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

  inspector.uiaFilterLabel = CreateWindowExW(0, WC_STATIC,
    newWideCString("UIA filter (type, name, AutomationId):"),
    WS_CHILD or WS_VISIBLE, 0, 0, 0, 0, inspector.hwnd, cast[HMENU](idUiaFilterLabel), hInst, nil)
  discard SendMessage(inspector.uiaFilterLabel, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.uiaFilterEdit = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, nil,
    WS_CHILD or WS_VISIBLE or WS_TABSTOP or ES_AUTOHSCROLL,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idUiaFilterEdit), hInst, nil)
  discard SendMessage(inspector.uiaFilterEdit, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.mainTree = CreateWindowExW(DWORD(WS_EX_CLIENTEDGE), WC_TREEVIEWW, nil,
    DWORD(treeStyle), 0, 0, 0, 0, inspector.hwnd, cast[HMENU](idMainTree), hInst, nil)
  discard SendMessage(inspector.mainTree, WM_SETFONT, WPARAM(font), LPARAM(TRUE))

  inspector.statusBar = CreateWindowExW(0, STATUSCLASSNAMEW, nil,
    WS_CHILD or WS_VISIBLE or SBARS_SIZEGRIP,
    0, 0, 0, 0, inspector.hwnd, cast[HMENU](idStatusBar), hInst, nil)
  discard SendMessage(inspector.statusBar, WM_SETFONT, WPARAM(font), LPARAM(TRUE))
  discard SendMessage(inspector.statusBar, SB_SIMPLE, WPARAM(FALSE), 0)

  discard EnableWindow(inspector.btnInvoke, FALSE)
  discard EnableWindow(inspector.btnFocus, FALSE)
  discard EnableWindow(inspector.btnHighlight, FALSE)
  discard EnableWindow(inspector.btnExpand, FALSE)

  resetWindowInfo(inspector)
  refreshWindowList(inspector)
  updateStatusBar(inspector)
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

proc followHighlightEnabled(inspector: InspectorWindow): bool =
  inspector.state.highlightFollow

proc updateFollowHighlight(inspector: InspectorWindow; enabled: bool) =
  inspector.state.highlightFollow = enabled
  if inspector.followHighlightCheck != 0:
    discard SendMessage(inspector.followHighlightCheck, BM_SETCHECK,
      WPARAM(if enabled: BST_CHECKED else: BST_UNCHECKED), 0)
  if inspector.hwnd != 0:
    if enabled:
      inspector.lastHoverHwnd = HWND(0)
      inspector.lastHoverTick = 0
      discard SetTimer(inspector.hwnd, followMouseTimerId, UINT(followMouseIntervalMs), nil)
    else:
      KillTimer(inspector.hwnd, followMouseTimerId)

proc autoHighlight(inspector: InspectorWindow; element: ptr IUIAutomationElement) =
  if not inspector.followHighlightEnabled():
    return
  if element.isNil:
    return
  discard highlightElementBounds(element, inspector.highlightColor, 1500, inspector.logger)

proc refreshCurrentRoot(inspector: InspectorWindow) =
  if inspector.uia.isNil:
    return
  let hwndSel = inspector.selectedWindowHandle()
  if hwndSel != HWND(0):
    try:
      let element = inspector.uia.fromWindowHandle(hwndSel)
      updateActionContext(inspector, element, some(ElementIdentifier(path: @[])))
      inspector.uia.setRootElement(element)
    except CatchableError as exc:
      if inspector.logger != nil:
        var fields = @[("error", exc.msg)]
        addContext(inspector, fields)
        inspector.logger.warn("Failed to refresh selected window root", fields)
      inspector.uia.setRootElement(nil)
  else:
    inspector.uia.setRootElement(nil)
  rebuildElementTree(inspector)
  let element = inspector.elementFromSelection()
  if not element.isNil:
    defer: discard element.Release()
    autoHighlight(inspector, element)

proc handleFollowMouseTimer(inspector: InspectorWindow) =
  if inspector.uia.isNil or not inspector.followHighlightEnabled():
    return
  var pt: POINT
  discard GetCursorPos(addr pt)
  let hwnd = WindowFromPoint(pt)
  if hwnd == 0:
    return
  if hwnd == inspector.hwnd or IsChild(inspector.hwnd, hwnd) != 0:
    return
  let now = GetTickCount()
  if hwnd == inspector.lastHoverHwnd and now - inspector.lastHoverTick < DWORD(followMouseDebounceMs):
    return

  inspector.lastHoverHwnd = hwnd
  inspector.lastHoverTick = now

  try:
    let element = inspector.uia.fromPoint(pt.x, pt.y)
    if not element.isNil:
      defer: discard element.Release()
      discard highlightElementBounds(element, inspector.highlightColor, 800, inspector.logger)
  except CatchableError as exc:
    if inspector.logger != nil:
      var fields = @[("error", exc.msg)]
      addContext(inspector, fields)
      inspector.logger.debug("Follow-highlight from mouse failed", fields)

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
    let sel = inspector.elementFromSelection()
    if not sel.isNil:
      defer: discard sel.Release()
      updateActionContext(inspector, sel, inspector.currentSelectionId())
    else:
      updateActionContext(inspector, nil, inspector.currentSelectionId())
    refreshWindowList(inspector)
  of idRefreshTree:
    let sel = inspector.elementFromSelection()
    if not sel.isNil:
      defer: discard sel.Release()
      updateActionContext(inspector, sel, inspector.currentSelectionId())
    refreshCurrentRoot(inspector)
  of idHighlightFollow:
    let enabled = SendMessage(inspector.followHighlightCheck, BM_GETCHECK, 0, 0) == BST_CHECKED
    updateFollowHighlight(inspector, enabled)
    saveInspectorState(inspector.statePath, inspector.state, inspector.logger)
    if enabled:
      let element = inspector.elementFromSelection()
      if not element.isNil:
        defer: discard element.Release()
        autoHighlight(inspector, element)
  of idUiaFilterEdit:
    if code == EN_CHANGE:
      rebuildElementTree(inspector)
  else:
    discard

proc handleNotify(inspector: InspectorWindow; lParam: LPARAM) =
  let hdr = cast[ptr NMHDR](lParam)
  if hdr.hwndFrom == inspector.mainTree and hdr.code == UINT(TVN_ITEMEXPANDINGW):
    try:
      let info = cast[ptr NMTREEVIEWW](lParam)
      if info.action == TVE_EXPAND:
        refreshTreeItemChildren(inspector, info.itemNew.hItem)
    except CatchableError as exc:
      if inspector.logger != nil:
        var fields = @[("error", exc.msg)]
        addContext(inspector, fields)
        inspector.logger.error("Tree expansion handling failed", fields)
  elif hdr.hwndFrom == inspector.mainTree and hdr.code == UINT(TVN_SELCHANGEDW):
    try:
      let info = cast[ptr NMTREEVIEWW](lParam)
      if info.itemNew.hItem in inspector.nodes:
        let element = inspector.elementFromId(inspector.nodes[info.itemNew.hItem])
        if not element.isNil:
          defer: discard element.Release()
          updateActionContext(inspector, element, some(inspector.nodes[info.itemNew.hItem]))
          autoHighlight(inspector, element)
          populateProperties(inspector, element)
        else:
          populateProperties(inspector, nil)
      else:
        populateProperties(inspector, nil)
    except CatchableError as exc:
      if inspector.logger != nil:
        var fields = @[("error", exc.msg)]
        addContext(inspector, fields)
        inspector.logger.error("Selection change handling failed", fields)
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
        let sel = inspector.elementFromSelection()
        if not sel.isNil:
          defer: discard sel.Release()
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
    try:
      let item = TreeView_GetSelection(inspector.patternsTree)
      if item != 0 and item in inspector.patternActions:
        let action = inspector.patternActions[item]
        let element = inspector.elementFromSelection()
        if not element.isNil:
          defer: discard element.Release()
          executePatternAction(inspector, element, action)
    except CatchableError as exc:
      if inspector.logger != nil:
        var fields = @[("error", exc.msg)]
        addContext(inspector, fields)
        inspector.logger.error("Pattern action failed", fields)
  elif hdr.hwndFrom == inspector.statusBar and (hdr.code == UINT(NM_CLICK) or hdr.code == UINT(NM_DBLCLK)):
    if inspector.accPath.len > 0:
      copyToClipboard(inspector.accPath, inspector.logger)

var inspectorClassRegistered = false

proc inspectorWndProc(hwnd: HWND; msg: UINT; wParam: WPARAM; lParam: LPARAM): LRESULT {.stdcall.} =
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
      try:
        registerMenu(inspector)
        createControls(inspector)
        inspectors[hwnd] = inspector
      except CatchableError as exc:
        if inspector.logger != nil:
          var fields = @[("error", exc.msg)]
          addContext(inspector, fields)
          inspector.logger.error("Inspector creation failed", fields)
        return -1
    return 0
  of UINT(msgInitTree):
    if inspector != nil:
      try:
        rebuildElementTree(inspector)
      except CatchableError as exc:
        if inspector.logger != nil:
          var fields = @[("error", exc.msg)]
          addContext(inspector, fields)
          inspector.logger.error("Failed to build UIA tree", fields)
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
  of WM_RBUTTONUP, WM_CONTEXTMENU:
    if inspector != nil and inspector.followHighlightEnabled():
      let ptOpt = screenPointFromMsg(inspector, hwnd, msg, lParam)
      if ptOpt.isSome:
        selectElementUnderCursor(inspector, ptOpt.get())
    return 0
  of WM_COMMAND:
    if inspector != nil:
      handleCommand(inspector, wParam, lParam)
    return 0
  of WM_NOTIFY:
    if inspector != nil:
      try:
        handleNotify(inspector, lParam)
      except CatchableError as exc:
        if inspector.logger != nil:
          var fields = @[("error", exc.msg)]
          addContext(inspector, fields)
          inspector.logger.error("Notification handling failed", fields)
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
    elif inspector != nil and UINT_PTR(wParam) == followMouseTimerId:
      handleFollowMouseTimer(inspector)
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
      KillTimer(hwnd, UINT_PTR(followMouseTimerId))
      if inspector.hwnd in inspectors:
        inspectors.del(inspector.hwnd)
    return 0
  else:
    discard
  result = DefWindowProc(hwnd, msg, wParam, lParam)

proc registerInspectorClass() =
  if inspectorClassRegistered:
    return
  var wc: WNDCLASSEXW
  wc.cbSize = UINT(sizeof(wc))
  wc.style = CS_HREDRAW or CS_VREDRAW
  wc.lpfnWndProc = inspectorWndProc
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
      logger.error("Failed to create inspector window",
        [("error", lastErrorMessage())])
    return false

  insp.hwnd = hwnd
  inspectors[hwnd] = insp
  updateFollowHighlight(insp, insp.state.highlightFollow)
  discard ShowWindow(hwnd, SW_SHOWNORMAL)
  discard UpdateWindow(hwnd)
  discard PostMessage(hwnd, UINT(msgInitTree), 0, 0)
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

  try:
    if not showInspectorWindow(uia, logger, resolvedStatePath):
      return 0
  except CatchableError as exc:
    if logger != nil:
      logger.error("Failed to start inspector window", [("error", exc.msg)])
    return 0

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
