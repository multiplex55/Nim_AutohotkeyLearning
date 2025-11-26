import std/[unittest]

when defined(windows):
  import uia

  suite "uia":
    test "init and root element":
      let automation = initUia()
      defer: automation.shutdown()
      check automation.rootElement() != nil

    test "findFirst returns nil for missing elements":
      let automation = initUia()
      defer: automation.shutdown()
      let missing = automation.findFirstByName("UIA-NON-EXISTENT", tsDescendants)
      check missing == nil

    test "fromPoint retrieves an element":
      let automation = initUia()
      defer: automation.shutdown()
      check automation.fromPoint(0, 0) != nil
else:
  static:
    echo "Skipping UIA tests on non-Windows platforms."
