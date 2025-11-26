import std/[monotimes, strformat, times]

when defined(windows):
  import input/input
  import uia
  import win/win

  type
    BenchmarkResult = object
      name: string
      iterations: int
      totalMicros: float
      avgMicros: float

  template runBench(name: string; iterations: int; body: untyped): BenchmarkResult =
    var start = getMonoTime()
    for _ in 0 ..< iterations:
      body
    let dur = getMonoTime() - start
    let totalMicros = dur.inNanoseconds.float / 1_000.0
    BenchmarkResult(
      name: name,
      iterations: iterations,
      totalMicros: totalMicros,
      avgMicros: totalMicros / iterations.float
    )

  proc printResult(res: BenchmarkResult) =
    echo fmt"{res.name}: {res.avgMicros:0.3f} µs avg over {res.iterations} iterations ({res.totalMicros / 1000:0.3f} ms total)"

  proc benchmarkInputDispatch(iterations = 500): BenchmarkResult =
    let noDelay = InputDelays(betweenEvents: 0.milliseconds, betweenChars: 0.milliseconds)
    runBench("SendInput dispatch (mouse move)", iterations):
      moveMouse(MousePoint(x: 0, y: 0), relative = true, delays = noDelay)

  proc benchmarkWindowEnumeration(iterations = 50): BenchmarkResult =
    runBench("Window enumeration", iterations):
      discard listWindows(includeUntitled = true)

  proc benchmarkUiaQuery(iterations = 100): BenchmarkResult =
    let automation = initUia()
    defer: automation.shutdown()
    runBench("UIA rootElement()", iterations):
      discard automation.rootElement()

  proc main() =
    echo "Hot-path microbenchmarks (lower is better):"
    echo "Run with -d:release for realistic timings."
    var results: seq[BenchmarkResult]
    results.add benchmarkInputDispatch()
    results.add benchmarkWindowEnumeration()
    results.add benchmarkUiaQuery()
    for res in results:
      printResult(res)

  when isMainModule:
    main()
else:
  static:
    echo "Hot-path microbenchmarks require Windows; skipping."
