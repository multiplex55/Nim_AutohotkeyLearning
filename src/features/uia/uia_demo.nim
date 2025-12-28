import std/[options, strformat, strutils, os]

import ../../core/logging
import ./uia
import ../../platform/windows/processes as winProcesses
import ../../platform/windows/windows as winWindows

import winim/lean
import winim/com
import winim/inc/uiautomation


proc ensureHrOk(hr: HRESULT, ctx: string) =
  if FAILED(hr):
    raise newException(UiaError, fmt"{ctx} failed (0x{hr:X})")

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

proc findNotepadWindow(): HWND =
  result = FindWindowW("Notepad", nil)
  if result != 0:
    return

  let untitled = winWindows.findWindowByTitleExact("Untitled - Notepad")
  if untitled != 0:
    result = HWND(untitled)

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

proc safeControlType(element: ptr IUIAutomationElement): int =
  try:
    element.currentControlType()
  except CatchableError:
    -1

proc safeNativeWindowHandle(element: ptr IUIAutomationElement): int =
  try:
    element.nativeWindowHandle()
  except CatchableError:
    0

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

proc formatElementInfo(element: ptr IUIAutomationElement): seq[(string, string)] =
  let name = safeCurrentName(element)
  let automationId = safeAutomationId(element)
  let runtimeId = safeRuntimeId(element)
  let bounds = safeBoundingRect(element)
  var fields: seq[(string, string)] = @[
    ("controlType", controlTypeName(safeControlType(element)))
  ]
  if name.len > 0:
    fields.add(("name", name))
  if automationId.len > 0:
    fields.add(("automationId", automationId))
  if runtimeId.len > 0:
    fields.add(("runtimeId", runtimeId))
  if bounds.isSome:
    let (left, top, width, height) = bounds.get()
    fields.add(("bounds", fmt"[{left.int}, {top.int}, {width.int}, {height.int}]"))

  fields

proc logElementTree(uia: Uia, element: ptr IUIAutomationElement, walker: ptr IUIAutomationTreeWalker, depth, maxDepth: int, logger: Logger) =
  if element.isNil or depth > maxDepth:
    return

  let hwnd = element.nativeWindowHandle()
  var fields: seq[(string, string)] = @[("depth", $depth)]
  fields.add(formatElementInfo(element))
  if hwnd != 0:
    fields.add(("hwnd", fmt"0x{cast[uint](hwnd):X}"))

  let indent = "  ".repeat(depth)
  logger.info(indent & "- UIA element", fields)

  if depth == maxDepth:
    return

  var child: ptr IUIAutomationElement
  let hrFirst = walker.GetFirstChildElement(element, addr child)
  if FAILED(hrFirst):
    ensureHrOk(hrFirst, "GetFirstChildElement")
  if hrFirst == S_FALSE or child.isNil:
    return

  var current = child
  while current != nil:
    logElementTree(uia, current, walker, depth + 1, maxDepth, logger)

    var next: ptr IUIAutomationElement
    let hrNext = walker.GetNextSiblingElement(current, addr next)
    discard current.Release()
    if FAILED(hrNext):
      ensureHrOk(hrNext, "GetNextSiblingElement")
    if hrNext == S_FALSE:
      break
    current = next

proc formatWindowInfo(hwnd: HWND): seq[(string, string)] =
  let hwndValue = cast[uint](hwnd)
  @[
    ("title", winWindows.getWindowTitle(hwnd)),
    ("hwnd", fmt"0x{hwndValue:X} ({hwndValue})")
  ]

proc matchesElement(element: ptr IUIAutomationElement, name, automationId: string, controlTypes: openArray[int]): bool =
  let ctrlType = safeControlType(element)
  let nameMatch = name.len == 0 or safeCurrentName(element) == name
  let idMatch = automationId.len == 0 or safeAutomationId(element) == automationId
  var typeMatch = controlTypes.len == 0
  if not typeMatch:
    for ct in controlTypes:
      if ct == ctrlType:
        typeMatch = true
        break
  nameMatch and idMatch and typeMatch

proc findElement(uia: Uia, element: ptr IUIAutomationElement, walker: ptr IUIAutomationTreeWalker, name, automationId: string, controlTypes: openArray[int], depth, maxDepth: int, logger: Logger): ptr IUIAutomationElement =
  if element.isNil or depth > maxDepth:
    return nil

  if matchesElement(element, name, automationId, controlTypes):
    discard element.AddRef()
    return element

  var child: ptr IUIAutomationElement
  let hrFirst = walker.GetFirstChildElement(element, addr child)
  if FAILED(hrFirst):
    logger.warn("Failed to enumerate children", [("depth", $depth), ("hresult", fmt"0x{hrFirst:X}")])
    return nil
  if hrFirst == S_FALSE or child.isNil:
    return nil

  var current = child
  while current != nil:
    let found = findElement(uia, current, walker, name, automationId, controlTypes, depth + 1, maxDepth, logger)
    if found != nil:
      discard current.Release()
      return found

    var next: ptr IUIAutomationElement
    let hrNext = walker.GetNextSiblingElement(current, addr next)
    discard current.Release()
    if FAILED(hrNext):
      logger.warn("Failed to enumerate siblings", [("depth", $depth), ("hresult", fmt"0x{hrNext:X}")])
      break
    if hrNext == S_FALSE:
      break
    current = next

  nil

proc appendBoolProperty(label: string, getter: proc(): bool, fields: var seq[(string, string)], logger: Logger) =
  try:
    let value = getter()
    fields.add((label, if value: "true" else: "false"))
  except CatchableError as exc:
    logger.warn(
      "Failed to read UIA property",
      [("property", label), ("error", exc.msg)]
    )

proc runUiaDemo*(maxDepth: int = 4): int =
  ## Windows-only UIA demo.
  ##
  ## Prerequisites: Windows host with Notepad available and UIA enabled. The
  ## demo launches (or reuses) Notepad, walks its UIA subtree, and logs a
  ## structured element outline (control type/name/automationId/runtimeId/hwnd)
  ## plus metadata about the edit control.
  var logger = newLogger()
  let existing = winProcesses.findProcessesByName("notepad.exe")
  if existing.len == 0:
    logger.info("Notepad not running; starting it")
    if not winProcesses.startProcessDetached("notepad.exe"):
      logger.error("Failed to start notepad.exe")
      return 1
    sleep(500)
  else:
    logger.info("Using existing Notepad instance", [("count", $existing.len)])

  var hwnd = findNotepadWindow()
  var attempts = 0
  while hwnd == 0 and attempts < 10:
    sleep(200)
    hwnd = findNotepadWindow()
    inc attempts

  if hwnd == 0:
    logger.error("Could not find a Notepad window. Is Notepad visible?")
    return 1

  logger.info(
    "Found Notepad window",
    formatWindowInfo(hwnd)
  )

  let uia = initUia()
  defer: uia.shutdown()

  let element = uia.fromWindowHandle(hwnd)
  if element.isNil:
    logger.error("Failed to obtain UIA element for Notepad window")
    return 1
  defer: discard element.Release()

  var walker: ptr IUIAutomationTreeWalker
  ensureHrOk(uia.automation.get_RawViewWalker(addr walker), "RawViewWalker")
  defer:
    if walker != nil:
      discard walker.Release()

  logger.info("Dumping Notepad UIA subtree", [("maxDepth", $maxDepth)])
  logElementTree(uia, element, walker, 0, maxDepth, logger)

  let editElement = findElement(
    uia,
    element,
    walker,
    name = "Text Editor",
    automationId = "15",
    controlTypes = [UIA_DocumentControlTypeId, UIA_EditControlTypeId],
    depth = 0,
    maxDepth = maxDepth + 5,
    logger = logger
  )

  if editElement.isNil:
    logger.error(
      "Could not locate Notepad edit control",
      [
        ("expectedName", "Text Editor"),
        ("automationId", "15"),
        ("controlTypes", "Document/Edit")
      ]
    )
    return 1

  defer: discard editElement.Release()

  let hwndVal = safeNativeWindowHandle(editElement)
  var elementFields = formatElementInfo(editElement)
  if hwndVal != 0:
    elementFields.add(("hwnd", fmt"0x{cast[uint](hwndVal):X}"))
  appendBoolProperty(
    "isEnabled",
    proc(): bool = editElement.isEnabled(),
    elementFields,
    logger
  )
  appendBoolProperty(
    "isKeyboardFocusable",
    proc(): bool = editElement.isKeyboardFocusable(),
    elementFields,
    logger
  )
  appendBoolProperty(
    "hasKeyboardFocus",
    proc(): bool = editElement.hasKeyboardFocus(),
    elementFields,
    logger
  )
  appendBoolProperty(
    "isContentElement",
    proc(): bool = editElement.isContentElement(),
    elementFields,
    logger
  )
  logger.info("Located Notepad edit control", elementFields)

  try:
    ensureHrOk(editElement.SetFocus(), "SetFocus(Edit)")
    logger.info("Set focus on Notepad edit control")
  except CatchableError as exc:
    logger.warn("Failed to set focus on edit control", [("error", exc.msg)])

  var valuePattern: ptr IUIAutomationValuePattern
  let hrValue = editElement.GetCurrentPattern(
    UIA_ValuePatternId,
    cast[ptr ptr IUnknown](addr valuePattern)
  )

  if SUCCEEDED(hrValue) and valuePattern != nil:
    var currentVal: BSTR
    if SUCCEEDED(valuePattern.get_CurrentValue(addr currentVal)):
      let textLength = if currentVal != nil: int(SysStringLen(currentVal)) else: 0
      logger.info("Retrieved edit control text", [("length", $textLength)])
      if currentVal != nil:
        SysFreeString(currentVal)
    else:
      logger.warn("Unable to retrieve edit control text via ValuePattern")
    discard valuePattern.Release()
  else:
    logger.warn(
      "ValuePattern not available for edit control",
      [("hresult", fmt"0x{hrValue:X}")]
    )

  result = 0
