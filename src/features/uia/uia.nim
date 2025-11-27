## Windows UI Automation helpers that mirror UIA-v2 style utilities.
import std/[times, os, strformat]

import winim/com
import winim/inc/objbase
import winim/inc/uiautomationclient
import winim/inc/windef

type
  UiaError* = object of CatchableError
  Uia* = ref object
    automation*: ptr IUIAutomation
    coInitialized: bool
    rootCache: ptr IUIAutomationElement
  TreeScope* = int32

const
  tsElement* = TreeScope(TreeScope_Element)
  tsChildren* = TreeScope(TreeScope_Children)
  tsDescendants* = TreeScope(TreeScope_Descendants)
  tsSubtree* = TreeScope(TreeScope_Subtree)

  # Property ids
  NamePropertyId* = UIA_NamePropertyId
  AutomationIdPropertyId* = UIA_AutomationIdPropertyId
  ClassNamePropertyId* = UIA_ClassNamePropertyId
  ControlTypePropertyId* = UIA_ControlTypePropertyId

proc checkHr(hr: HRESULT, ctx: string) =
  if FAILED(hr):
    raise newException(UiaError, fmt"{ctx} failed (0x{hr:X})")

proc initUia*(coInit: DWORD = COINIT_APARTMENTTHREADED): Uia =
  ## Initialize COM and create an IUIAutomation instance.
  let hr = CoInitializeEx(nil, coInit)
  if hr != S_OK and hr != S_FALSE and hr != RPC_E_CHANGED_MODE:
    raise newException(UiaError, fmt"COM initialization failed (0x{hr:X})")

  var automation: ptr IUIAutomation
  var coStarted = hr == S_OK or hr == S_FALSE

  try:
    checkHr(CoCreateInstance(CLSID_CUIAutomation, nil, DWORD(CLSCTX_INPROC_SERVER),
      IID_IUIAutomation, cast[ptr pointer](addr automation)), "CoCreateInstance(IUIAutomation)")

    result = Uia(automation: automation, coInitialized: coStarted)
  except:
    if automation != nil:
      discard automation.Release()
    if coStarted:
      CoUninitialize()
    raise

proc shutdown*(uia: Uia) =
  ## Release the UIA object and uninitialize COM if we started it.
  if uia.isNil: return
  if uia.rootCache != nil:
    discard uia.rootCache.Release()
    uia.rootCache = nil
  if uia.automation != nil:
    discard uia.automation.Release()
    uia.automation = nil
  if uia.coInitialized:
    CoUninitialize()

proc rootElement*(uia: Uia): ptr IUIAutomationElement =
  if uia.rootCache != nil:
    return uia.rootCache

  var element: ptr IUIAutomationElement
  checkHr(uia.automation.GetRootElement(addr element), "GetRootElement")
  uia.rootCache = element
  result = element

proc fromPoint*(uia: Uia, x, y: int32): ptr IUIAutomationElement =
  var element: ptr IUIAutomationElement
  var pt = POINT(x: x, y: y)
  checkHr(uia.automation.ElementFromPoint(pt, addr element), "ElementFromPoint")
  result = element

proc fromWindowHandle*(uia: Uia, hwnd: HWND): ptr IUIAutomationElement =
  ## Retrieve an element for a given window handle.
  var element: ptr IUIAutomationElement
  checkHr(uia.automation.ElementFromHandle(UIA_HWND(hwnd), addr element), "ElementFromHandle")
  result = element

proc currentName*(element: ptr IUIAutomationElement): string =
  var text: BSTR
  checkHr(element.get_CurrentName(addr text), "get_CurrentName")
  result = $text

proc currentClassName*(element: ptr IUIAutomationElement): string =
  var text: BSTR
  checkHr(element.get_CurrentClassName(addr text), "get_CurrentClassName")
  result = $text

proc currentAutomationId*(element: ptr IUIAutomationElement): string =
  var text: BSTR
  checkHr(element.get_CurrentAutomationId(addr text), "get_CurrentAutomationId")
  result = $text

proc currentControlType*(element: ptr IUIAutomationElement): int =
  var controlType: cint
  checkHr(element.get_CurrentControlType(addr controlType), "get_CurrentControlType")
  result = controlType

proc variantFromString(value: string): VARIANT =
  VariantInit(addr result)
  result.vt = VT_BSTR
  result.bstrVal = SysAllocString(value)

proc variantFromInt(value: int): VARIANT =
  VariantInit(addr result)
  result.vt = VT_I4
  result.lVal = LONG(value)

proc propertyCondition*(uia: Uia, propId: int, value: VARIANT): ptr IUIAutomationCondition =
  var cond: ptr IUIAutomationCondition
  checkHr(uia.automation.CreatePropertyCondition(PROPERTYID(propId), value, addr cond), "CreatePropertyCondition")
  discard VariantClear(addr value)
  result = cond

proc nameCondition*(uia: Uia, value: string): ptr IUIAutomationCondition =
  result = propertyCondition(uia, NamePropertyId, variantFromString(value))

proc automationIdCondition*(uia: Uia, value: string): ptr IUIAutomationCondition =
  result = propertyCondition(uia, AutomationIdPropertyId, variantFromString(value))

proc classNameCondition*(uia: Uia, value: string): ptr IUIAutomationCondition =
  result = propertyCondition(uia, ClassNamePropertyId, variantFromString(value))

proc controlTypeCondition*(uia: Uia, value: int): ptr IUIAutomationCondition =
  result = propertyCondition(uia, ControlTypePropertyId, variantFromInt(value))

proc andCondition*(uia: Uia, lhs, rhs: ptr IUIAutomationCondition): ptr IUIAutomationCondition =
  var cond: ptr IUIAutomationCondition
  checkHr(uia.automation.CreateAndCondition(lhs, rhs, addr cond), "CreateAndCondition")
  result = cond

proc nameAndControlType*(uia: Uia, name: string, controlTypeId: int): ptr IUIAutomationCondition =
  ## Combine a name + control type into a single condition.
  result = uia.andCondition(uia.nameCondition(name), uia.controlTypeCondition(controlTypeId))

proc automationIdAndControlType*(uia: Uia, automationId: string, controlTypeId: int): ptr IUIAutomationCondition =
  ## Combine an automation id + control type into a single condition.
  result = uia.andCondition(uia.automationIdCondition(automationId), uia.controlTypeCondition(controlTypeId))

proc findFirst*(uia: Uia, scope: TreeScope, cond: ptr IUIAutomationCondition,
                root: ptr IUIAutomationElement = nil): ptr IUIAutomationElement =
  var start = root
  if start.isNil:
    start = uia.rootElement()
  var found: ptr IUIAutomationElement
  let hr = start.FindFirst(scope, cond, addr found)
  if hr == UIA_E_ELEMENTNOTAVAILABLE:
    return nil
  checkHr(hr, "FindFirst")
  result = found

proc findAll*(uia: Uia, scope: TreeScope, cond: ptr IUIAutomationCondition,
              root: ptr IUIAutomationElement = nil): seq[ptr IUIAutomationElement] =
  var start = root
  if start.isNil:
    start = uia.rootElement()
  var arr: ptr IUIAutomationElementArray
  let hr = start.FindAll(scope, cond, addr arr)
  if hr == UIA_E_ELEMENTNOTAVAILABLE:
    return @[]
  checkHr(hr, "FindAll")

  var length: cint
  checkHr(arr.get_Length(addr length), "ElementArray.get_Length")
  for i in 0 ..< length:
    var el: ptr IUIAutomationElement
    checkHr(arr.GetElement(i, addr el), "ElementArray.GetElement")
    result.add el

proc waitElement*(uia: Uia, scope: TreeScope, cond: ptr IUIAutomationCondition,
                  timeout: Duration = initDuration(seconds = 3), pollInterval = initDuration(milliseconds = 250),
                  root: ptr IUIAutomationElement = nil): ptr IUIAutomationElement =
  let deadline = times.now() + timeout
  var currentRoot = root
  while times.now() < deadline:
    let found = uia.findFirst(scope, cond, currentRoot)
    if not found.isNil:
      return found
    sleep(pollInterval.inMilliseconds())
  return nil

proc hasPattern*(element: ptr IUIAutomationElement, patternId: int, patternName: string): bool =
  ## Check if a UIA pattern is available on the element.
  var obj: pointer
  let hr = element.TryGetCurrentPattern(patternId, addr obj)
  if hr == S_OK and obj != nil:
    discard cast[ptr IUnknown](obj).Release()
    return true
  if hr == UIA_E_ELEMENTNOTAVAILABLE:
    return false
  return false

proc availablePatterns*(element: ptr IUIAutomationElement): seq[string] =
  ## Return the set of known patterns exposed by the element.
  const knownPatterns = [
    (UIA_InvokePatternId, "Invoke"),
    (UIA_ValuePatternId, "Value"),
    (UIA_RangeValuePatternId, "RangeValue"),
    (UIA_SelectionItemPatternId, "SelectionItem"),
    (UIA_SelectionPatternId, "Selection"),
    (UIA_TogglePatternId, "Toggle"),
    (UIA_ExpandCollapsePatternId, "ExpandCollapse"),
    (UIA_WindowPatternId, "Window"),
    (UIA_ScrollPatternId, "Scroll"),
    (UIA_GridPatternId, "Grid"),
    (UIA_GridItemPatternId, "GridItem"),
    (UIA_TextPatternId, "Text"),
    (UIA_TablePatternId, "Table"),
    (UIA_TableItemPatternId, "TableItem"),
    (UIA_DockPatternId, "Dock"),
    (UIA_LegacyIAccessiblePatternId, "LegacyIAccessible"),
    (UIA_ScrollItemPatternId, "ScrollItem"),
    (UIA_TransformPatternId, "Transform"),
    (UIA_ItemContainerPatternId, "ItemContainer"),
    (UIA_AnnotationPatternId, "Annotation"),
    (UIA_SpreadsheetPatternId, "Spreadsheet"),
    (UIA_SpreadsheetItemPatternId, "SpreadsheetItem"),
    (UIA_StylesPatternId, "Styles"),
    (UIA_DragPatternId, "Drag"),
    (UIA_DropTargetPatternId, "DropTarget"),
    (UIA_TextChildPatternId, "TextChild"),
    (UIA_TextEditPatternId, "TextEdit"),
    (UIA_CustomNavigationPatternId, "CustomNavigation")
  ]

  for (patternId, patternName) in knownPatterns:
    if element.hasPattern(patternId, patternName):
      result.add patternName

proc controlTypeName*(controlType: int): string =
  ## Convert a control type id into a friendly name.
  case controlType
  of UIA_ButtonControlTypeId: "Button"
  of UIA_CalendarControlTypeId: "Calendar"
  of UIA_CheckBoxControlTypeId: "CheckBox"
  of UIA_ComboBoxControlTypeId: "ComboBox"
  of UIA_EditControlTypeId: "Edit"
  of UIA_HyperlinkControlTypeId: "Hyperlink"
  of UIA_ImageControlTypeId: "Image"
  of UIA_ListItemControlTypeId: "ListItem"
  of UIA_ListControlTypeId: "List"
  of UIA_MenuControlTypeId: "Menu"
  of UIA_MenuBarControlTypeId: "MenuBar"
  of UIA_MenuItemControlTypeId: "MenuItem"
  of UIA_ProgressBarControlTypeId: "ProgressBar"
  of UIA_RadioButtonControlTypeId: "RadioButton"
  of UIA_ScrollBarControlTypeId: "ScrollBar"
  of UIA_SliderControlTypeId: "Slider"
  of UIA_SpinnerControlTypeId: "Spinner"
  of UIA_StatusBarControlTypeId: "StatusBar"
  of UIA_TabControlTypeId: "Tab"
  of UIA_TabItemControlTypeId: "TabItem"
  of UIA_TextControlTypeId: "Text"
  of UIA_ToolBarControlTypeId: "ToolBar"
  of UIA_ToolTipControlTypeId: "ToolTip"
  of UIA_TreeControlTypeId: "Tree"
  of UIA_TreeItemControlTypeId: "TreeItem"
  of UIA_CustomControlTypeId: "Custom"
  of UIA_GroupControlTypeId: "Group"
  of UIA_ThumbControlTypeId: "Thumb"
  of UIA_DataGridControlTypeId: "DataGrid"
  of UIA_DataItemControlTypeId: "DataItem"
  of UIA_DocumentControlTypeId: "Document"
  of UIA_SplitButtonControlTypeId: "SplitButton"
  of UIA_WindowControlTypeId: "Window"
  of UIA_PaneControlTypeId: "Pane"
  of UIA_HeaderControlTypeId: "Header"
  of UIA_HeaderItemControlTypeId: "HeaderItem"
  of UIA_TableControlTypeId: "Table"
  of UIA_TitleBarControlTypeId: "TitleBar"
  of UIA_SeparatorControlTypeId: "Separator"
  of UIA_SemanticZoomControlTypeId: "SemanticZoom"
  of UIA_AppBarControlTypeId: "AppBar"
  else: "Unknown"

proc nativeWindowHandle*(element: ptr IUIAutomationElement): int =
  ## Retrieve the native window handle associated with an element.
  var hwnd: cint
  checkHr(element.get_CurrentNativeWindowHandle(addr hwnd), "get_CurrentNativeWindowHandle")
  result = hwnd

proc hasKeyboardFocus*(element: ptr IUIAutomationElement): bool =
  var focused: BOOL
  checkHr(element.get_CurrentHasKeyboardFocus(addr focused), "get_CurrentHasKeyboardFocus")
  result = focused != 0

proc isKeyboardFocusable*(element: ptr IUIAutomationElement): bool =
  var focusable: BOOL
  checkHr(element.get_CurrentIsKeyboardFocusable(addr focusable), "get_CurrentIsKeyboardFocusable")
  result = focusable != 0

proc isControlElement*(element: ptr IUIAutomationElement): bool =
  var controlElem: BOOL
  checkHr(element.get_CurrentIsControlElement(addr controlElem), "get_CurrentIsControlElement")
  result = controlElem != 0

proc isContentElement*(element: ptr IUIAutomationElement): bool =
  var contentElem: BOOL
  checkHr(element.get_CurrentIsContentElement(addr contentElem), "get_CurrentIsContentElement")
  result = contentElem != 0

proc isPassword*(element: ptr IUIAutomationElement): bool =
  var password: BOOL
  checkHr(element.get_CurrentIsPassword(addr password), "get_CurrentIsPassword")
  result = password != 0

proc requirePattern[T](element: ptr IUIAutomationElement, patternId: int, patternName: string): ptr T =
  var obj: pointer
  let hr = element.GetCurrentPattern(patternId, addr obj)
  if FAILED(hr) or obj.isNil:
    raise newException(UiaError, fmt"Pattern '{patternName}' not available (0x{hr:X})")
  result = cast[ptr T](obj)

proc invoke*(element: ptr IUIAutomationElement) =
  let pattern = element.requirePattern[IUIAutomationInvokePattern](UIA_InvokePatternId, "Invoke")
  checkHr(pattern.Invoke(), "Invoke")

proc setValue*(element: ptr IUIAutomationElement, value: string) =
  let pattern = element.requirePattern[IUIAutomationValuePattern](UIA_ValuePatternId, "Value")
  checkHr(pattern.SetValue(value), "Value.SetValue")

proc select*(element: ptr IUIAutomationElement) =
  let pattern = element.requirePattern[IUIAutomationSelectionItemPattern](UIA_SelectionItemPatternId, "SelectionItem")
  checkHr(pattern.Select(), "SelectionItem.Select")

proc addToSelection*(element: ptr IUIAutomationElement) =
  let pattern = element.requirePattern[IUIAutomationSelectionItemPattern](UIA_SelectionItemPatternId, "SelectionItem")
  checkHr(pattern.AddToSelection(), "SelectionItem.AddToSelection")

proc removeFromSelection*(element: ptr IUIAutomationElement) =
  let pattern = element.requirePattern[IUIAutomationSelectionItemPattern](UIA_SelectionItemPatternId, "SelectionItem")
  checkHr(pattern.RemoveFromSelection(), "SelectionItem.RemoveFromSelection")

proc toggle*(element: ptr IUIAutomationElement) =
  let pattern = element.requirePattern[IUIAutomationTogglePattern](UIA_TogglePatternId, "Toggle")
  checkHr(pattern.Toggle(), "Toggle.Toggle")

proc expand*(element: ptr IUIAutomationElement) =
  let pattern = element.requirePattern[IUIAutomationExpandCollapsePattern](UIA_ExpandCollapsePatternId, "ExpandCollapse")
  checkHr(pattern.Expand(), "ExpandCollapse.Expand")

proc collapse*(element: ptr IUIAutomationElement) =
  let pattern = element.requirePattern[IUIAutomationExpandCollapsePattern](UIA_ExpandCollapsePatternId, "ExpandCollapse")
  checkHr(pattern.Collapse(), "ExpandCollapse.Collapse")

proc scroll*(element: ptr IUIAutomationElement, horizontalAmount, verticalAmount: double) =
  let pattern = element.requirePattern[IUIAutomationScrollPattern](UIA_ScrollPatternId, "Scroll")
  checkHr(pattern.Scroll(horizontalAmount, verticalAmount), "Scroll.Scroll")

proc legacyDoDefaultAction*(element: ptr IUIAutomationElement) =
  let pattern = element.requirePattern[IUIAutomationLegacyIAccessiblePattern](UIA_LegacyIAccessiblePatternId, "LegacyIAccessible")
  checkHr(pattern.DoDefaultAction(), "Legacy.DoDefaultAction")

proc legacySetValue*(element: ptr IUIAutomationElement, value: string) =
  let pattern = element.requirePattern[IUIAutomationLegacyIAccessiblePattern](UIA_LegacyIAccessiblePatternId, "LegacyIAccessible")
  checkHr(pattern.SetValue(value), "Legacy.SetValue")

proc findFirstByName*(uia: Uia, name: string, scope: TreeScope = tsDescendants, root: ptr IUIAutomationElement = nil): ptr IUIAutomationElement =
  uia.findFirst(scope, uia.nameCondition(name), root)

proc findFirstByAutomationId*(uia: Uia, automationId: string, scope: TreeScope = tsDescendants, root: ptr IUIAutomationElement = nil): ptr IUIAutomationElement =
  uia.findFirst(scope, uia.automationIdCondition(automationId), root)

proc findButtonByName*(uia: Uia, name: string, scope: TreeScope = tsDescendants, root: ptr IUIAutomationElement = nil): ptr IUIAutomationElement =
  uia.findFirst(scope, uia.nameAndControlType(name, UIA_ButtonControlTypeId), root)

proc findButtonByAutomationId*(uia: Uia, automationId: string, scope: TreeScope = tsDescendants, root: ptr IUIAutomationElement = nil): ptr IUIAutomationElement =
  uia.findFirst(scope, uia.automationIdAndControlType(automationId, UIA_ButtonControlTypeId), root)

proc findWindowButtonByName*(uia: Uia, window: ptr IUIAutomationElement, name: string): ptr IUIAutomationElement =
  ## Find the first descendant button within a window by its name.
  if window.isNil:
    return nil
  result = uia.findFirst(tsDescendants, uia.nameAndControlType(name, UIA_ButtonControlTypeId), window)

proc findWindowButtonByAutomationId*(uia: Uia, window: ptr IUIAutomationElement, automationId: string): ptr IUIAutomationElement =
  ## Find the first descendant button within a window by its automation id.
  if window.isNil:
    return nil
  result = uia.findFirst(tsDescendants, uia.automationIdAndControlType(automationId, UIA_ButtonControlTypeId), window)

proc findWindowButtonByControlType*(uia: Uia, window: ptr IUIAutomationElement, controlTypeId: int): ptr IUIAutomationElement =
  ## Find the first descendant control of the provided type within a window.
  if window.isNil:
    return nil
  result = uia.findFirst(tsDescendants, uia.controlTypeCondition(controlTypeId), window)

proc isEnabled*(element: ptr IUIAutomationElement): bool =
  var enabled: BOOL
  checkHr(element.get_CurrentIsEnabled(addr enabled), "get_CurrentIsEnabled")
  result = enabled != 0

proc isOffscreen*(element: ptr IUIAutomationElement): bool =
  var offscreen: BOOL
  checkHr(element.get_CurrentIsOffscreen(addr offscreen), "get_CurrentIsOffscreen")
  result = offscreen != 0

proc isVisible*(element: ptr IUIAutomationElement): bool =
  result = not element.isOffscreen()

proc windowPattern*(element: ptr IUIAutomationElement): ptr IUIAutomationWindowPattern =
  result = element.requirePattern[IUIAutomationWindowPattern](UIA_WindowPatternId, "Window")

proc windowVisualState*(element: ptr IUIAutomationElement): WindowVisualState =
  var state: WindowVisualState
  checkHr(element.windowPattern().get_CurrentWindowVisualState(addr state), "Window.get_CurrentWindowVisualState")
  result = state

proc closeWindow*(element: ptr IUIAutomationElement) =
  ## Close a window element using the Window pattern.
  checkHr(element.windowPattern().Close(), "Window.Close")

proc isMinimized*(element: ptr IUIAutomationElement): bool =
  element.windowVisualState() == WindowVisualState_Minimized

proc isMaximized*(element: ptr IUIAutomationElement): bool =
  element.windowVisualState() == WindowVisualState_Maximized

proc isNormal*(element: ptr IUIAutomationElement): bool =
  element.windowVisualState() == WindowVisualState_Normal
