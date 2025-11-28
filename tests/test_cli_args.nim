import std/[options, unittest]

import ../src/cli/arguments

suite "cli args":
  test "uia demo flag is opt-in":
    let parsed = parseCliArgs(@["--uia-demo"])
    check parsed.uiaDemo
    check parsed.configPath.isNone
    check parsed.uiaMaxDepth == 4

  test "custom config path without demo flag":
    let parsed = parseCliArgs(@["custom.toml"])
    check parsed.uiaDemo == false
    check parsed.configPath.get("missing") == "custom.toml"
