import std/[options, strformat, strutils, tables]

import ../../core/logging
import ./uia
import ./uia_plugin

import winim/lean
import winim/com
import winim/inc/uiautomation
import wNim

when not defined(windows):
  {.fatal: "UI Automation inspector only runs on Windows".}

type
  UiaTreeNode = ref object
    element: ptr IUIAutomationElement
    label: string
    children: seq[UiaTreeNode]

  TreeFilters = object
    name: string
    automationId: string
    controlType: string

proc safeGetProperty(element: ptr IUIAutomationElement, propertyId: PROPERTYID): Option[VARIANT] =
  try:
    var value: VARIANT
    checkHr(element.GetCurrentPropertyValue(propertyId, addr value), "GetCurrentPropertyValue")
    return some(value)
  except CatchableError:
    return none(VARIANT)

proc safeString(element: ptr IUIAutomationElement, propertyId: PROPERTYID): string =
  let maybeVal = safeGetProperty(element, propertyId)
  if maybeVal.isSome:
    var mutableVal = maybeVal.get()
    defer: discard VariantClear(addr mutableVal)
    result = $mutableVal.bstrVal
  else:
    result = ""

proc safeControlType(element: ptr IUIAutomationElement): int =
  let maybeVal = safeGetProperty(element, UIA_ControlTypePropertyId)
  if maybeVal.isSome:
    var mutableVal = maybeVal.get()
    defer: discard VariantClear(addr mutableVal)
    result = int(mutableVal.lVal)
  else:
    result = -1

proc safeRuntimeId(element: ptr IUIAutomationElement): string =
  var arr: ptr SAFEARRAY
  let hr = element.GetRuntimeId(addr arr)
  if FAILED(hr) or arr.isNil:
    return ""
  defer: discard SafeArrayDestroy(arr)

  var lbound, ubound: LONG
  if FAILED(SafeArrayGetLBound(arr, 1, addr lbound)) or FAILED(SafeArrayGetUBound(arr, 1, addr ubound)):
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

proc safeBounds(element: ptr IUIAutomationElement): string =
  var rectVar: VARIANT
  let hr = element.GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId, addr rectVar)
  defer: discard VariantClear(addr rectVar)
  if FAILED(hr) or rectVar.parray.isNil or (rectVar.vt and VT_ARRAY) == 0:
    return ""

  var lbound, ubound: LONG
  if FAILED(SafeArrayGetLBound(rectVar.parray, 1, addr lbound)) or FAILED(SafeArrayGetUBound(rectVar.parray, 1, addr ubound)):
    return ""
  if ubound - lbound + 1 < 4:
    return ""

  var coords: array[4, float64]
  var idx = lbound
  var i = 0
  while i < 4:
    if FAILED(SafeArrayGetElement(rectVar.parray, addr idx, addr coords[i])):
      return ""
    inc idx
    inc i

  fmt"[{coords[0].int}, {coords[1].int}, {coords[2].int}, {coords[3].int}]"

proc describeNode(element: ptr IUIAutomationElement): string =
  let name = safeString(element, UIA_NamePropertyId)
  let automationId = safeString(element, UIA_AutomationIdPropertyId)
  let controlType = controlTypeName(safeControlType(element))

  var pieces: seq[string] = @[controlType]
  if name.len > 0:
    pieces.add("name=\"" & name & "\"")
  if automationId.len > 0:
    pieces.add("automationId=\"" & automationId & "\"")
  pieces.join(" | ")

proc buildTree(element: ptr IUIAutomationElement, walker: ptr IUIAutomationTreeWalker, depth, maxDepth: int): UiaTreeNode =
  if element.isNil or depth > maxDepth:
    return nil

  result = UiaTreeNode(element: element, label: describeNode(element))
  var child: ptr IUIAutomationElement
  let hrFirst = walker.GetFirstChildElement(element, addr child)
  if FAILED(hrFirst) or child.isNil:
    return

  var current = child
  while current != nil:
    let childNode = buildTree(current, walker, depth + 1, maxDepth)
    if childNode != nil:
      result.children.add(childNode)

    var next: ptr IUIAutomationElement
    let hrNext = walker.GetNextSiblingElement(current, addr next)
    if FAILED(hrNext) or hrNext == S_FALSE:
      break
    current = next

proc releaseTree(node: UiaTreeNode) =
  if node.isNil:
    return
  for child in node.children:
    releaseTree(child)
  discard node.element.Release()

proc matchesFilter(node: UiaTreeNode, filters: TreeFilters): bool =
  if node.isNil:
    return false

  let name = safeString(node.element, UIA_NamePropertyId).toLowerAscii()
  let automationId = safeString(node.element, UIA_AutomationIdPropertyId).toLowerAscii()
  let controlType = controlTypeName(safeControlType(node.element)).toLowerAscii()

  let nameOk = filters.name.len == 0 or name.contains(filters.name.toLowerAscii())
  let idOk = filters.automationId.len == 0 or automationId.contains(filters.automationId.toLowerAscii())
  let typeOk = filters.controlType.len == 0 or controlType.contains(filters.controlType.toLowerAscii())

  result = nameOk and idOk and typeOk

proc filterTree(node: UiaTreeNode, filters: TreeFilters): UiaTreeNode =
  if node.isNil:
    return nil

  var keptChildren: seq[UiaTreeNode] = @[]
  for child in node.children:
    let filtered = filterTree(child, filters)
    if filtered != nil:
      keptChildren.add(filtered)

  if matchesFilter(node, filters) or keptChildren.len > 0:
    result = UiaTreeNode(element: node.element, label: node.label, children: keptChildren)

proc addToTree(tree: TreeCtrl, node: UiaTreeNode, parent: wTreeItem, index: var Table[wTreeItem, UiaTreeNode]) =
  if node.isNil:
    return

  let item =
    if parent.isOk:
      tree.appendItem(parent, node.label)
    else:
      tree.addRoot(node.label)

  index[item] = node
  for child in node.children:
    addToTree(tree, child, item, index)

proc hasPattern(element: ptr IUIAutomationElement, patternId: PATTERNID): bool =
  var obj: ptr IUnknown
  let hr = element.GetCurrentPattern(patternId, addr obj)
  if SUCCEEDED(hr) and obj != nil:
    discard obj.Release()
    return true
  false

proc updatePropertyList(list: ListCtrl, node: UiaTreeNode) =
  list.deleteAllItems()
  if node.isNil:
    return

  let element = node.element
  let entries = [
    ("Name", safeString(element, UIA_NamePropertyId)),
    ("AutomationId", safeString(element, UIA_AutomationIdPropertyId)),
    ("ClassName", safeString(element, UIA_ClassNamePropertyId)),
    ("ControlType", controlTypeName(safeControlType(element))),
    ("RuntimeId", safeRuntimeId(element)),
    ("Bounds", safeBounds(element))
  ]

  for idx, (key, value) in entries:
    discard list.insertItem(idx, key)
    list.setItem(idx, 1, value)

proc toggleTree(tree: TreeCtrl, item: wTreeItem, expand: bool) =
  if not item.isOk:
    return
  var child = tree.getFirstChild(item)
  while child.isOk:
    toggleTree(tree, child, expand)
    child = tree.getNextSibling(child)

  if expand:
    item.expand()
  else:
    item.collapse()

when isMainModule:
  proc main() =
    let logger = newLogger()
    let uiaClient = initUia()
    defer: uiaClient.shutdown()

    let rootElement = uiaClient.rootElement()
    if rootElement.isNil:
      echo "UIA returned nil root element"
      quit(1)

    var walker: ptr IUIAutomationTreeWalker
    checkHr(uiaClient.automation.get_RawViewWalker(addr walker), "RawViewWalker")
    defer: discard walker.Release()

    let rootModel = buildTree(rootElement, walker, 0, 4)
    var activeModel = rootModel
    var filters = TreeFilters()

    var nodeIndex = initTable[wTreeItem, UiaTreeNode]()

    let app = App()
    let frame = Frame(title = "UIA Inspector", size = (1100, 700))

    let leftPanel = Panel(frame)
    let rightPanel = Panel(frame)

    let nameLabel = StaticText(leftPanel, label = "Name:")
    let nameFilter = TextCtrl(leftPanel, value = "", style = wTeProcessTab)
    let automationLabel = StaticText(leftPanel, label = "AutomationId:")
    let automationFilter = TextCtrl(leftPanel, value = "", style = wTeProcessTab)
    let controlTypeLabel = StaticText(leftPanel, label = "ControlType:")
    let controlTypeFilter = TextCtrl(leftPanel, value = "", style = wTeProcessTab)
    let applyFilterBtn = Button(leftPanel, label = "Apply Filters")
    let clearFilterBtn = Button(leftPanel, label = "Clear")
    let expandAllBtn = Button(leftPanel, label = "Expand All")
    let collapseAllBtn = Button(leftPanel, label = "Collapse All")

    let treeCtrl = TreeCtrl(leftPanel)

    let propertyList = ListCtrl(rightPanel, style = wLcReport or wLcSingleSel)
    discard propertyList.insertColumn(0, "Property", width = 150)
    discard propertyList.insertColumn(1, "Value", width = 400)

    let invokeBtn = Button(rightPanel, label = "Invoke")
    let focusBtn = Button(rightPanel, label = "Set Focus")
    let closeBtn = Button(rightPanel, label = "Close")

    frame.layout:
      leftPanel:
        top = frame.top
        left = frame.left
        bottom = frame.bottom
        width = 520
      rightPanel:
        top = frame.top
        left = leftPanel.right
        right = frame.right
        bottom = frame.bottom

    let pad = 8

    leftPanel.layout:
      nameLabel:
        top = leftPanel.top + pad
        left = leftPanel.left + pad
      nameFilter:
        centerY = nameLabel.centerY
        left = nameLabel.right + pad
        width = 200
      automationLabel:
        centerY = nameLabel.centerY
        left = nameFilter.right + pad
      automationFilter:
        centerY = nameLabel.centerY
        left = automationLabel.right + pad
        width = 200
      controlTypeLabel:
        centerY = nameLabel.centerY
        left = automationFilter.right + pad
      controlTypeFilter:
        centerY = nameLabel.centerY
        left = controlTypeLabel.right + pad
        width = 160
      applyFilterBtn:
        centerY = nameLabel.centerY
        left = controlTypeFilter.right + pad
      clearFilterBtn:
        centerY = nameLabel.centerY
        left = applyFilterBtn.right + pad
      expandAllBtn:
        centerY = nameLabel.centerY
        left = clearFilterBtn.right + pad
      collapseAllBtn:
        centerY = nameLabel.centerY
        left = expandAllBtn.right + pad
      treeCtrl:
        top = nameLabel.bottom + pad
        left = leftPanel.left + pad
        right = leftPanel.right - pad
        bottom = leftPanel.bottom - pad

    rightPanel.layout:
      propertyList:
        top = rightPanel.top + pad
        left = rightPanel.left + pad
        right = rightPanel.right - pad
        bottom = invokeBtn.top - pad
      invokeBtn:
        left = rightPanel.left + pad
        bottom = rightPanel.bottom - pad
      focusBtn:
        centerY = invokeBtn.centerY
        left = invokeBtn.right + pad
      closeBtn:
        centerY = invokeBtn.centerY
        left = focusBtn.right + pad

    var selectedNode: UiaTreeNode

    proc rebuildTree() =
      treeCtrl.deleteAllItems()
      nodeIndex.clear()
      if activeModel != nil:
        addToTree(treeCtrl, activeModel, wTreeItem(), nodeIndex)
        let rootItem = treeCtrl.getRootItem()
        if rootItem.isOk:
          rootItem.expand()
      selectedNode = nil
      updatePropertyList(propertyList, selectedNode)
      invokeBtn.disable()
      focusBtn.disable()
      closeBtn.disable()

    proc syncFilters() =
      filters.name = nameFilter.getValue()
      filters.automationId = automationFilter.getValue()
      filters.controlType = controlTypeFilter.getValue()

    proc applyFilters() =
      syncFilters()
      if filters.name.len == 0 and filters.automationId.len == 0 and filters.controlType.len == 0:
        activeModel = rootModel
      else:
        activeModel = filterTree(rootModel, filters)
      rebuildTree()

    nameFilter.connect(wEvent_TextEnter) do (e: wEvent):
      applyFilters()
    automationFilter.connect(wEvent_TextEnter) do (e: wEvent):
      applyFilters()
    controlTypeFilter.connect(wEvent_TextEnter) do (e: wEvent):
      applyFilters()
    applyFilterBtn.connect(wEvent_Button) do (e: wEvent):
      applyFilters()

    clearFilterBtn.connect(wEvent_Button) do (e: wEvent):
      filters = TreeFilters()
      nameFilter.setValue("")
      automationFilter.setValue("")
      controlTypeFilter.setValue("")
      activeModel = rootModel
      rebuildTree()

    expandAllBtn.connect(wEvent_Button) do (e: wEvent):
      let rootItem = treeCtrl.getRootItem()
      toggleTree(treeCtrl, rootItem, true)

    collapseAllBtn.connect(wEvent_Button) do (e: wEvent):
      let rootItem = treeCtrl.getRootItem()
      toggleTree(treeCtrl, rootItem, false)

    treeCtrl.connect(wEvent_TreeSelChanged) do (e: wEvent):
      let item = e.getItem()
      if nodeIndex.hasKey(item):
        selectedNode = nodeIndex[item]
        updatePropertyList(propertyList, selectedNode)
        invokeBtn.enable(hasPattern(selectedNode.element, UIA_InvokePatternId))
        closeBtn.enable(hasPattern(selectedNode.element, UIA_WindowPatternId))
        focusBtn.enable(not selectedNode.isNil)
      else:
        selectedNode = nil
        updatePropertyList(propertyList, selectedNode)
        invokeBtn.disable()
        closeBtn.disable()
        focusBtn.disable()

    invokeBtn.connect(wEvent_Button) do (e: wEvent):
      if not selectedNode.isNil and hasPattern(selectedNode.element, UIA_InvokePatternId):
        try:
          uiaClient.invoke(selectedNode.element)
        except CatchableError as exc:
          logger.error("Invoke failed", [("error", exc.msg)])

    focusBtn.connect(wEvent_Button) do (e: wEvent):
      if not selectedNode.isNil:
        try:
          checkHr(selectedNode.element.SetFocus(), "SetFocus")
        except CatchableError as exc:
          logger.error("SetFocus failed", [("error", exc.msg)])

    closeBtn.connect(wEvent_Button) do (e: wEvent):
      if not selectedNode.isNil and hasPattern(selectedNode.element, UIA_WindowPatternId):
        try:
          uiaClient.closeWindow(selectedNode.element)
        except CatchableError as exc:
          logger.error("Close failed", [("error", exc.msg)])

    rebuildTree()
    invokeBtn.disable()
    focusBtn.disable()
    closeBtn.disable()

    frame.center()
    frame.show()
    discard app.mainLoop()

    releaseTree(rootModel)

  main()
