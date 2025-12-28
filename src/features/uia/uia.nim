## Windows UI Automation helpers with graceful fallback when UIA headers are unavailable.

import std/[options, strformat]

import winim/com
import winim/inc/objbase
import winim/inc/windef
import winim/inc/uiautomation

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
    let hrCreate = CoCreateInstance(
      addr CLSID_CUIAutomation,         # REFCLSID  (ptr GUID)
      nil,
      CLSCTX_INPROC_SERVER,             # DWORD, same value but clearer
      addr IID_IUIAutomation,           # REFIID    (ptr GUID)
      cast[ptr LPVOID](addr automation) # ptr LPVOID (ptr pointer)
    )

    checkHr(hrCreate, "CoCreateInstance")
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

proc setRootElement*(uia: Uia, element: ptr IUIAutomationElement) =
  ## Replace the cached root element, releasing the previous reference if present.
  if uia.rootCache != nil:
    discard uia.rootCache.Release()
  uia.rootCache = element

proc fromPoint*(uia: Uia, x, y: int32): ptr IUIAutomationElement =
  var element: ptr IUIAutomationElement
  var pt = POINT(x: x, y: y)
  checkHr(uia.automation.ElementFromPoint(pt, addr element), "ElementFromPoint")
  result = element


proc fromWindowHandle*(uia: Uia, hwnd: HWND): ptr IUIAutomationElement =
  var element: ptr IUIAutomationElement

  let uiaHwnd: UIA_HWND = cast[UIA_HWND](hwnd)

  checkHr(
    uia.automation.ElementFromHandle(uiaHwnd, addr element),
    "ElementFromHandle"
  )
  result = element


proc getCurrentPropertyValue*(uia: Uia, element: ptr IUIAutomationElement,
    propertyId: PROPERTYID): VARIANT =
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
  checkHr(element.GetCurrentPropertyValue(UIA_IsKeyboardFocusablePropertyId,
      addr val), "IsKeyboardFocusable")
  result = asBool(val)

proc hasKeyboardFocus*(element: ptr IUIAutomationElement): bool =
  var val: VARIANT
  checkHr(element.GetCurrentPropertyValue(UIA_HasKeyboardFocusPropertyId,
      addr val), "HasKeyboardFocus")
  result = asBool(val)

proc isControlElement*(element: ptr IUIAutomationElement): bool =
  var val: VARIANT
  checkHr(element.GetCurrentPropertyValue(UIA_IsControlElementPropertyId,
      addr val), "IsControlElement")
  result = asBool(val)

proc isContentElement*(element: ptr IUIAutomationElement): bool =
  var val: VARIANT
  checkHr(element.GetCurrentPropertyValue(UIA_IsContentElementPropertyId,
      addr val), "IsContentElement")
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

proc safeBoundingRect*(element: ptr IUIAutomationElement): Option[(float, float,
    float, float)] =
  ## Best-effort retrieval of the element's bounding rectangle.
  var rectVar: VARIANT
  let hr = element.GetCurrentPropertyValue(
    UIA_BoundingRectanglePropertyId,
    addr rectVar
  )
  defer:
    discard VariantClear(addr rectVar)

  if FAILED(hr) or rectVar.parray.isNil or (rectVar.vt and VT_ARRAY) == 0:
    return

  var lbound, ubound: LONG
  if FAILED(SafeArrayGetLBound(rectVar.parray, 1, addr lbound)) or
      FAILED(SafeArrayGetUBound(rectVar.parray, 1, addr ubound)):
    return
  if ubound - lbound + 1 < 4:
    return

  var coords: array[4, float64]
  var idx = lbound
  var i = 0
  while i < 4:
    if FAILED(SafeArrayGetElement(rectVar.parray, addr idx, addr coords[i])):
      return
    inc idx
    inc i

  some((coords[0].float, coords[1].float, coords[2].float, coords[3].float))

proc availablePatterns*(element: ptr IUIAutomationElement): seq[string] =
  ## Stub: GetSupportedPatterns is not available in the current winim UIA bindings.
  ## We don't rely on this for core automation, so return an empty list for now.
  result = @[]


proc invoke*(uia: Uia, element: ptr IUIAutomationElement) =
  var pattern: ptr IUIAutomationInvokePattern
  checkHr(
    element.GetCurrentPattern(
      UIA_InvokePatternId,
      cast[ptr ptr IUnknown](addr pattern)
    ),
    "GetCurrentPattern(Invoke)"
  )
  checkHr(pattern.Invoke(), "Invoke")

proc closeWindow*(uia: Uia, element: ptr IUIAutomationElement) =
  var pattern: ptr IUIAutomationWindowPattern
  checkHr(
    element.GetCurrentPattern(
      UIA_WindowPatternId,
      cast[ptr ptr IUnknown](addr pattern)
    ),
    "GetCurrentPattern(Window)"
  )
  discard pattern.Close()


proc windowVisualState*(element: ptr IUIAutomationElement): int =
  var pattern: ptr IUIAutomationWindowPattern
  checkHr(
    element.GetCurrentPattern(
      UIA_WindowPatternId,
      cast[ptr ptr IUnknown](addr pattern)
    ),
    "GetCurrentPattern(Window)"
  )

  var state: WindowVisualState
  checkHr(pattern.get_CurrentWindowVisualState(addr state), "CurrentWindowVisualState")
  result = int(state)

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
