## Windows UI Automation helpers with graceful fallback when UIA headers are unavailable.
import std/[times, os, strformat]

const uiaHeadersAvailable* = compiles(import winim/inc/uiautomationclient)

when uiaHeadersAvailable:
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
    var element: ptr IUIAutomationElement
    checkHr(uia.automation.ElementFromHandle(UIA_HWND(hwnd), addr element), "ElementFromHandle")
    result = element

  proc propertyValue*(uia: Uia, element: ptr IUIAutomationElement, propertyId: int): VARIANT =
    var val: VARIANT
    checkHr(element.GetCurrentPropertyValue(propertyId, addr val), "GetCurrentPropertyValue")
    result = val

  proc asBool(v: VARIANT): bool =
    v.vt == VT_BOOL and v.boolVal == VARIANT_TRUE

  proc isEnabled*(element: ptr IUIAutomationElement): bool =
    var val: VARIANT
    checkHr(element.GetCurrentPropertyValue(UIA_IsEnabledPropertyId, addr val), "IsEnabled")
    result = asBool(val)

  proc isOffscreen*(element: ptr IUIAutomationElement): bool =
    var val: VARIANT
    checkHr(element.GetCurrentPropertyValue(UIA_IsOffscreenPropertyId, addr val), "IsOffscreen")
    result = asBool(val)

  proc isVisible*(element: ptr IUIAutomationElement): bool =
    not element.isOffscreen()

  proc isKeyboardFocusable*(element: ptr IUIAutomationElement): bool =
    var val: VARIANT
    checkHr(element.GetCurrentPropertyValue(UIA_IsKeyboardFocusablePropertyId, addr val), "IsKeyboardFocusable")
    result = asBool(val)

  proc hasKeyboardFocus*(element: ptr IUIAutomationElement): bool =
    var val: VARIANT
    checkHr(element.GetCurrentPropertyValue(UIA_HasKeyboardFocusPropertyId, addr val), "HasKeyboardFocus")
    result = asBool(val)

  proc isControlElement*(element: ptr IUIAutomationElement): bool =
    var val: VARIANT
    checkHr(element.GetCurrentPropertyValue(UIA_IsControlElementPropertyId, addr val), "IsControlElement")
    result = asBool(val)

  proc isContentElement*(element: ptr IUIAutomationElement): bool =
    var val: VARIANT
    checkHr(element.GetCurrentPropertyValue(UIA_IsContentElementPropertyId, addr val), "IsContentElement")
    result = asBool(val)

  proc isPassword*(element: ptr IUIAutomationElement): bool =
    var val: VARIANT
    checkHr(element.GetCurrentPropertyValue(UIA_IsPasswordPropertyId, addr val), "IsPassword")
    result = asBool(val)

  proc currentName*(element: ptr IUIAutomationElement): string =
    var val: VARIANT
    checkHr(element.GetCurrentPropertyValue(UIA_NamePropertyId, addr val), "CurrentName")
    $val.bstrVal

  proc currentAutomationId*(element: ptr IUIAutomationElement): string =
    var val: VARIANT
    checkHr(element.GetCurrentPropertyValue(UIA_AutomationIdPropertyId, addr val), "CurrentAutomationId")
    $val.bstrVal

  proc currentClassName*(element: ptr IUIAutomationElement): string =
    var val: VARIANT
    checkHr(element.GetCurrentPropertyValue(UIA_ClassNamePropertyId, addr val), "CurrentClassName")
    $val.bstrVal

  proc currentControlType*(element: ptr IUIAutomationElement): int =
    var val: VARIANT
    checkHr(element.GetCurrentPropertyValue(UIA_ControlTypePropertyId, addr val), "CurrentControlType")
    int(val.lVal)

  proc availablePatterns*(element: ptr IUIAutomationElement): seq[string] =
    var patIds: ptr SAFEARRAY
    checkHr(element.GetSupportedPatterns(addr patIds), "GetSupportedPatterns")
    if patIds == nil:
      return @[]
    let c = patIds.cElems
    setLen(result, int(c))
    var vals = cast[ptr int32](patIds.pvData)
    for i in 0 ..< int(c):
      result[i] = $vals[i]

  proc invoke*(uia: Uia, element: ptr IUIAutomationElement) =
    var pattern: ptr IUIAutomationInvokePattern
    checkHr(element.GetCurrentPattern(UIA_InvokePatternId, cast[ptr pointer](addr pattern)), "GetCurrentPattern(Invoke)")
    checkHr(pattern.Invoke(), "Invoke")

  proc closeWindow*(uia: Uia, element: ptr IUIAutomationElement) =
    var pattern: ptr IUIAutomationWindowPattern
    checkHr(element.GetCurrentPattern(UIA_WindowPatternId, cast[ptr pointer](addr pattern)), "GetCurrentPattern(Window)")
    discard pattern.Close()

  proc windowVisualState*(element: ptr IUIAutomationElement): int =
    var pattern: ptr IUIAutomationWindowPattern
    checkHr(element.GetCurrentPattern(UIA_WindowPatternId, cast[ptr pointer](addr pattern)), "GetCurrentPattern(Window)")
    var state: int
    checkHr(pattern.get_CurrentWindowVisualState(addr state), "CurrentWindowVisualState")
    state

  proc describeWindow*(uia: Uia, hwnd: HWND): string =
    var title = ""
    var className = ""
    var buffer = newStringOfCap(256)
    buffer.setLen(256)
    let copied = GetWindowText(hwnd, buffer.cstring, 256)
    if copied > 0:
      buffer.setLen(copied)
      title = buffer
    buffer.setLen(256)
    let copiedClass = GetClassName(hwnd, buffer.cstring, 256)
    if copiedClass > 0:
      buffer.setLen(copiedClass)
      className = buffer
    fmt"{title} ({className})"

  proc nativeWindowHandle*(element: ptr IUIAutomationElement): int =
    var hwnd: UIA_HWND
    checkHr(element.get_CurrentNativeWindowHandle(addr hwnd), "CurrentNativeWindowHandle")
    cast[int](hwnd)

else:
  import std/options
  type
    UiaError* = object of CatchableError
    Uia* = ref object
    TreeScope* = int32

  const
    tsElement* = TreeScope(0)
    tsChildren* = TreeScope(0)
    tsDescendants* = TreeScope(0)
    tsSubtree* = TreeScope(0)

  proc initUia*(coInit: int = 0): Uia =
    raise newException(UiaError, "UI Automation headers not available; install winim with UIA support")

  proc shutdown*(uia: Uia) = discard
  proc rootElement*(uia: Uia): pointer =
    raise newException(UiaError, "UI Automation unavailable")
  proc fromPoint*(uia: Uia, x, y: int32): pointer =
    raise newException(UiaError, "UI Automation unavailable")
  proc fromWindowHandle*(uia: Uia, hwnd: int): pointer =
    raise newException(UiaError, "UI Automation unavailable")

