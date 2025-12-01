## notepad_button_template.nim
##
## Uses project internals to:
##   1. Check if "Untitled - Notepad" exists (exact title).
##   2. Attach UIA and list all Button controls.
##   3. Run a hard-coded button click sequence using flexible matching:
##        - name (UIA Name)
##        - automationId
##        - className
##        - visibleOnly
##
## No hotkeys. No config. Runs once and exits.

when system.hostOS != "windows":
  {.error: "Windows-only example.".}

import std/[os, strformat]

import winim/lean
import winim/com
import winim/inc/uiautomation

import ./core/logging
import ./features/uia/uia as uiaLib
import ./platform/windows/windows as winWindows

# ─────────────────────────────────────────────────────────────────────────────
# Config / types
# ─────────────────────────────────────────────────────────────────────────────

const
  TargetWindowTitle = "Untitled - Notepad"  ## exact match
  MaxTreeDepth      = 8                     ## UIA depth for listing/search

type
  ButtonStep = object
    ## Any of these can be left empty to ignore that filter.
    name: string          ## UIA Name (e.g. "File")
    automationId: string  ## e.g. "ContentButton"
    className: string     ## e.g. "Button"
    visibleOnly: bool     ## if true, require uia.isVisible(element)
    delayMsAfter: int     ## milliseconds to sleep after click

## EXAMPLE: Click the "File" button
##
## From your log:
##   name=File, automationId=ContentButton, className=Button, visible=true
##
## You can add more steps later (Edit, View, etc).
let ButtonSequence: seq[ButtonStep] = @[
  ButtonStep(
    name: "File",
    automationId: "ContentButton",
    className: "Button",
    visibleOnly: true,
    delayMsAfter: 0
  ),
  ButtonStep(
    name: "Add New Tab",
    automationId: "AddButton",
    className: "Button",
    visibleOnly: true,
    delayMsAfter: 0
  ),
]

# ─────────────────────────────────────────────────────────────────────────────
# UIA traversal helpers (using your uiaLib)
# ─────────────────────────────────────────────────────────────────────────────

proc dumpButtonsSub(
    uia: uiaLib.Uia,
    element: ptr IUIAutomationElement,
    walker: ptr IUIAutomationTreeWalker,
    depth, maxDepth: int,
    logger: Logger
) =
  ## Recursive DFS that logs every Button control it finds.
  if element.isNil or depth > maxDepth:
    return

  try:
    let ctrlType = uiaLib.currentControlType(element)
    if ctrlType == UIA_ButtonControlTypeId:
      let name      = uiaLib.currentName(element)
      let autoId    = uiaLib.currentAutomationId(element)
      let className = uiaLib.currentClassName(element)
      let visible   = uiaLib.isVisible(element)

      if logger != nil:
        logger.info("Found UIA button",
          [ ("depth", $depth)
          , ("name", name)
          , ("automationId", autoId)
          , ("className", className)
          , ("visible", $visible)
          ])
  except CatchableError as exc:
    if logger != nil:
      logger.warn("Failed to read button properties",
        [("depth", $depth), ("error", exc.msg)])

  if depth == maxDepth:
    return

  var child: ptr IUIAutomationElement
  let hrFirst = walker.GetFirstChildElement(element, addr child)
  if FAILED(hrFirst) or child.isNil:
    if FAILED(hrFirst) and logger != nil:
      logger.warn("GetFirstChildElement failed",
        [("depth", $depth), ("hresult", fmt"0x{hrFirst:X}")])
    return

  var current = child
  while current != nil:
    dumpButtonsSub(uia, current, walker, depth + 1, maxDepth, logger)

    var next: ptr IUIAutomationElement
    let hrNext = walker.GetNextSiblingElement(current, addr next)
    discard current.Release()
    if FAILED(hrNext):
      if logger != nil:
        logger.warn("GetNextSiblingElement failed",
          [("depth", $depth), ("hresult", fmt"0x{hrNext:X}")])
      break
    if hrNext == S_FALSE or next.isNil:
      break
    current = next

proc listButtonsFromRoot(uia: uiaLib.Uia,
                         root: ptr IUIAutomationElement,
                         maxDepth: int,
                         logger: Logger) =
  ## Create a RawViewWalker and list all Button controls under root.
  var walker: ptr IUIAutomationTreeWalker
  let hrWalker = uia.automation.get_RawViewWalker(addr walker)
  if FAILED(hrWalker) or walker.isNil:
    if logger != nil:
      logger.error("Failed to create RawViewWalker",
        [("hresult", fmt"0x{hrWalker:X}")])
    return

  defer:
    discard walker.Release()

  if logger != nil:
    logger.info("Listing UIA buttons", [("maxDepth", $maxDepth)])

  dumpButtonsSub(uia, root, walker, 0, maxDepth, logger)

# ─────────────────────────────────────────────────────────────────────────────
# UIA search + click using flexible matching
# ─────────────────────────────────────────────────────────────────────────────

proc matchesButton(
    uia: uiaLib.Uia,
    element: ptr IUIAutomationElement,
    step: ButtonStep
): bool =
  ## Returns true if this element matches the ButtonStep selectors.
  try:
    if uiaLib.currentControlType(element) != UIA_ButtonControlTypeId:
      return false

    if step.name.len > 0 and uiaLib.currentName(element) != step.name:
      return false

    if step.automationId.len > 0 and
       uiaLib.currentAutomationId(element) != step.automationId:
      return false

    if step.className.len > 0 and
       uiaLib.currentClassName(element) != step.className:
      return false

    if step.visibleOnly and not uiaLib.isVisible(element):
      return false

    true
  except CatchableError:
    false

proc findButtonSub(
    uia: uiaLib.Uia,
    element: ptr IUIAutomationElement,
    walker: ptr IUIAutomationTreeWalker,
    step: ButtonStep,
    depth, maxDepth: int,
    logger: Logger
): ptr IUIAutomationElement =
  ## DFS that returns the first Button matching the ButtonStep filters.
  if element.isNil or depth > maxDepth:
    return nil

  if matchesButton(uia, element, step):
    discard element.AddRef()
    return element

  if depth == maxDepth:
    return nil

  var child: ptr IUIAutomationElement
  let hrFirst = walker.GetFirstChildElement(element, addr child)
  if FAILED(hrFirst) or child.isNil:
    if FAILED(hrFirst) and logger != nil:
      logger.warn("GetFirstChildElement failed while searching",
        [("depth", $depth), ("hresult", fmt"0x{hrFirst:X}")])
    return nil

  var current = child
  while current != nil:
    let found = findButtonSub(uia, current, walker, step, depth + 1, maxDepth, logger)
    if found != nil:
      discard current.Release()
      return found

    var next: ptr IUIAutomationElement
    let hrNext = walker.GetNextSiblingElement(current, addr next)
    discard current.Release()
    if FAILED(hrNext):
      if logger != nil:
        logger.warn("GetNextSiblingElement failed while searching",
          [("depth", $depth), ("hresult", fmt"0x{hrNext:X}")])
      break
    if hrNext == S_FALSE or next.isNil:
      break
    current = next

  nil

proc findButtonForStep(
    uia: uiaLib.Uia,
    root: ptr IUIAutomationElement,
    step: ButtonStep,
    maxDepth: int,
    logger: Logger
): ptr IUIAutomationElement =
  ## Locate a Button that matches the ButtonStep filters under root.
  if root.isNil:
    if logger != nil:
      logger.error("UIA root element is nil; cannot search for button")
    return nil

  var walker: ptr IUIAutomationTreeWalker
  let hrWalker = uia.automation.get_RawViewWalker(addr walker)
  if FAILED(hrWalker) or walker.isNil:
    if logger != nil:
      logger.error("Failed to create RawViewWalker for search",
        [("hresult", fmt"0x{hrWalker:X}")])
    return nil

  defer:
    discard walker.Release()

  result = findButtonSub(uia, root, walker, step, 0, maxDepth, logger)

proc runButtonSequence(
    uia: uiaLib.Uia,
    root: ptr IUIAutomationElement,
    logger: Logger
): bool =
  ## Run the hard-coded ButtonSequence.
  ##
  ## If a button is missing or a click fails, log and return false.
  if ButtonSequence.len == 0:
    if logger != nil:
      logger.info("No ButtonSequence configured; skipping clicks")
    return true

  for i, step in ButtonSequence:
    if logger != nil:
      logger.info("Sequence step",
        [("index", $i),
         ("name", step.name),
         ("automationId", step.automationId),
         ("className", step.className)])

    let el = findButtonForStep(uia, root, step, MaxTreeDepth, logger)
    if el.isNil:
      if logger != nil:
        logger.error("Sequence aborted – button not found",
          [("index", $i),
           ("name", step.name),
           ("automationId", step.automationId),
           ("className", step.className)])
      return false

    try:
      uiaLib.invoke(uia, el)
      if logger != nil:
        logger.info("Invoked button",
          [("name", uiaLib.currentName(el)),
           ("automationId", uiaLib.currentAutomationId(el)),
           ("className", uiaLib.currentClassName(el))])
    except CatchableError as exc:
      if logger != nil:
        logger.error("UIA invoke failed",
          [("error", exc.msg)])
      discard el.Release()
      return false

    discard el.Release()

    if step.delayMsAfter > 0:
      if logger != nil:
        logger.info("Sleeping after click",
          [("milliseconds", $step.delayMsAfter)])
      sleep(step.delayMsAfter)

  true

# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

proc main(): int =
  let logger = newLogger()
  logger.info("Starting Notepad button template",
    [("targetTitle", TargetWindowTitle)])

  # 1) Find the window by exact title using your windows helper.
  let hwnd = winWindows.findWindowByTitleExact(TargetWindowTitle)
  if hwnd == 0:
    logger.info("Target window not found; exiting")
    return 0

  logger.info("Found target window",
    [("hwnd", fmt"0x{cast[uint](hwnd):X}"),
     ("desc", winWindows.describeWindow(hwnd))])

  # 2) Initialize UIA via your wrapper.
  var uia: uiaLib.Uia
  try:
    uia = uiaLib.initUia()
  except CatchableError as exc:
    logger.error("Failed to initialize UI Automation", [("error", exc.msg)])
    return 1

  defer:
    if not uia.isNil:
      uia.shutdown()

  # 3) Attach to that window and get the root element.
  var root: ptr IUIAutomationElement
  try:
    root = uia.fromWindowHandle(HWND(hwnd))
  except CatchableError as exc:
    logger.error("fromWindowHandle failed", [("error", exc.msg)])
    return 1

  if root.isNil:
    logger.error("UIA returned nil root element")
    return 1

  # 4) List all Button controls (for discovery).
  listButtonsFromRoot(uia, root, MaxTreeDepth, logger)

  # 5) Run the hard-coded click sequence (File button, etc.).
  if not runButtonSequence(uia, root, logger):
    logger.error("Button sequence failed – exiting with error")
    return 1

  logger.info("Template finished successfully")
  return 0

when isMainModule:
  quit(main())
