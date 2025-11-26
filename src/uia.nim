## Windows UI Automation helpers that mirror UIA-v2 style utilities.
import std/[options, times, os, strformat]

when defined(windows):
  import winim/com
  import winim/inc/oleauto
  import winim/inc/objbase
  import winim/inc/uiautomationclient

  type
    UiaError* = object of CatchableError
    Uia* = ref object
      automation*: ptr IUIAutomation
      coInitialized: bool
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
    checkHr(CoCreateInstance(CLSID_CUIAutomation, nil, CLSCTX_INPROC_SERVER,
      IID_IUIAutomation, cast[ptr pointer](addr automation)), "CoCreateInstance(IUIAutomation)")

    result = Uia(automation: automation, coInitialized: hr == S_OK or hr == S_FALSE)

  proc shutdown*(uia: Uia) =
    ## Release the UIA object and uninitialize COM if we started it.
    if uia.isNil: return
    if uia.automation != nil:
      discard uia.automation.Release()
      uia.automation = nil
    if uia.coInitialized:
      CoUninitialize()

  proc rootElement*(uia: Uia): ptr IUIAutomationElement =
    var element: ptr IUIAutomationElement
    checkHr(uia.automation.GetRootElement(addr element), "GetRootElement")
    result = element

  proc fromPoint*(uia: Uia, x, y: int32): ptr IUIAutomationElement =
    var element: ptr IUIAutomationElement
    var pt = tagPOINT(x: x, y: y)
    checkHr(uia.automation.ElementFromPoint(pt, addr element), "ElementFromPoint")
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
    result.lVal = value

  proc propertyCondition*(uia: Uia, propId: int, value: VARIANT): ptr IUIAutomationCondition =
    var cond: ptr IUIAutomationCondition
    checkHr(uia.automation.CreatePropertyCondition(propId, value, addr cond), "CreatePropertyCondition")
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
                    timeout: Duration = 3.seconds, pollInterval = 250.milliseconds,
                    root: ptr IUIAutomationElement = nil): ptr IUIAutomationElement =
    let deadline = times.now() + timeout
    var currentRoot = root
    while times.now() < deadline:
      let found = uia.findFirst(scope, cond, currentRoot)
      if not found.isNil:
        return found
      sleep(pollInterval.inMilliseconds)
    return nil

  proc hasPattern*(element: ptr IUIAutomationElement, patternId: int, patternName: string): bool =
    discard patternName
    var obj: pointer
    let hr = element.TryGetCurrentPattern(patternId, addr obj)
    if hr == S_OK and obj != nil:
      return true
    if hr == UIA_E_ELEMENTNOTAVAILABLE:
      return false
    return false

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
    checkHr(pattern.DoDefaultAction(), "LegacyIAccessible.DoDefaultAction")

  proc legacySetValue*(element: ptr IUIAutomationElement, value: string) =
    let pattern = element.requirePattern[IUIAutomationLegacyIAccessiblePattern](UIA_LegacyIAccessiblePatternId, "LegacyIAccessible")
    checkHr(pattern.SetValue(value), "LegacyIAccessible.SetValue")

  proc findFirstByName*(uia: Uia, name: string, scope: TreeScope = tsDescendants,
                        root: ptr IUIAutomationElement = nil): ptr IUIAutomationElement =
    result = uia.findFirst(scope, uia.nameCondition(name), root)

  proc findFirstByAutomationId*(uia: Uia, automationId: string, scope: TreeScope = tsDescendants,
                                root: ptr IUIAutomationElement = nil): ptr IUIAutomationElement =
    result = uia.findFirst(scope, uia.automationIdCondition(automationId), root)

  proc findButtonByName*(uia: Uia, name: string, scope: TreeScope = tsDescendants,
                         root: ptr IUIAutomationElement = nil): ptr IUIAutomationElement =
    result = uia.findFirst(scope, uia.nameAndControlType(name, UIA_ButtonControlTypeId), root)

  proc findButtonByAutomationId*(uia: Uia, automationId: string, scope: TreeScope = tsDescendants,
                                 root: ptr IUIAutomationElement = nil): ptr IUIAutomationElement =
    result = uia.findFirst(scope, uia.automationIdAndControlType(automationId, UIA_ButtonControlTypeId), root)

else:
  type
    UiaError* = object of CatchableError
    Uia* = ref object
    TreeScope* = int32

  const
    tsElement* = TreeScope(0)
    tsChildren* = TreeScope(0)
    tsDescendants* = TreeScope(0)
    tsSubtree* = TreeScope(0)
    NamePropertyId* = 0
    AutomationIdPropertyId* = 0
    ClassNamePropertyId* = 0
    ControlTypePropertyId* = 0

  proc initUia*(coInit: DWORD = 0): Uia =
    raise newException(UiaError, "UI Automation requires Windows")

  template notWindows(): untyped =
    raise newException(UiaError, "UI Automation is only available on Windows.")

  proc shutdown*(uia: Uia) = notWindows()
  proc rootElement*(uia: Uia): pointer = notWindows()
  proc fromPoint*(uia: Uia, x, y: int32): pointer = notWindows()
  proc nameCondition*(uia: Uia, value: string): pointer = notWindows()
  proc automationIdCondition*(uia: Uia, value: string): pointer = notWindows()
  proc classNameCondition*(uia: Uia, value: string): pointer = notWindows()
  proc controlTypeCondition*(uia: Uia, value: int): pointer = notWindows()
  proc andCondition*(uia: Uia, lhs, rhs: pointer): pointer = notWindows()
  proc findFirst*(uia: Uia, scope: TreeScope, cond: pointer, root: pointer = nil): pointer = notWindows()
  proc findAll*(uia: Uia, scope: TreeScope, cond: pointer, root: pointer = nil): seq[pointer] = notWindows()
  proc waitElement*(uia: Uia, scope: TreeScope, cond: pointer, timeout: Duration = 3.seconds, pollInterval = 250.milliseconds, root: pointer = nil): pointer = notWindows()
  proc hasPattern*(element: pointer, patternId: int, patternName: string): bool = notWindows()
  proc invoke*(element: pointer) = notWindows()
  proc setValue*(element: pointer, value: string) = notWindows()
  proc select*(element: pointer) = notWindows()
  proc addToSelection*(element: pointer) = notWindows()
  proc removeFromSelection*(element: pointer) = notWindows()
  proc toggle*(element: pointer) = notWindows()
  proc expand*(element: pointer) = notWindows()
  proc collapse*(element: pointer) = notWindows()
  proc scroll*(element: pointer, horizontalAmount, verticalAmount: double) = notWindows()
  proc legacyDoDefaultAction*(element: pointer) = notWindows()
  proc legacySetValue*(element: pointer, value: string) = notWindows()
  proc findFirstByName*(uia: Uia, name: string, scope: TreeScope = tsDescendants, root: pointer = nil): pointer = notWindows()
  proc findFirstByAutomationId*(uia: Uia, automationId: string, scope: TreeScope = tsDescendants, root: pointer = nil): pointer = notWindows()
  proc findButtonByName*(uia: Uia, name: string, scope: TreeScope = tsDescendants, root: pointer = nil): pointer = notWindows()
  proc findButtonByAutomationId*(uia: Uia, automationId: string, scope: TreeScope = tsDescendants, root: pointer = nil): pointer = notWindows()
