import std/[monotimes, options, times]

import ./logging

type
  TaskId* = int

  TaskAction* = proc() {.closure.}

  TaskKind = enum
    tkOnce, tkRepeat, tkSequence

  ScheduledStep* = object
    delay*: Duration
    action*: TaskAction

  ScheduledTask = ref object
    id: TaskId
    kind: TaskKind
    nextRun: MonoTime
    interval: Duration
    action: TaskAction
    steps: seq[ScheduledStep]
    stepIndex: int
    cancelled: bool

  TaskHandle* = ref object
    id*: TaskId
    cancelled*: bool

  Scheduler* = ref object
    tasks: seq[ScheduledTask]
    nextId: TaskId
    logger*: Logger

proc newScheduler*(logger: Logger = nil): Scheduler =
  Scheduler(tasks: @[], nextId: 1, logger: logger)

proc cancel*(scheduler: Scheduler, handle: TaskHandle) =
  if scheduler == nil or handle == nil:
    return
  for task in scheduler.tasks.mitems:
    if task.id == handle.id:
      task.cancelled = true
      handle.cancelled = true
      if scheduler.logger != nil:
        scheduler.logger.debug("Cancelled scheduled task", [("id", $handle.id)])
      break

proc scheduleOnce*(scheduler: Scheduler, delay: Duration, action: TaskAction): TaskHandle =
  if scheduler == nil:
    return nil
  scheduler.nextId.inc
  let task = ScheduledTask(
    id: scheduler.nextId - 1,
    kind: tkOnce,
    nextRun: getMonoTime() + delay,
    interval: delay,
    action: action,
    steps: @[]
  )
  scheduler.tasks.add(task)
  if scheduler.logger != nil:
    scheduler.logger.debug("Scheduled one-shot task", [("id", $task.id), ("delay", $delay.inMilliseconds)])
  TaskHandle(id: task.id, cancelled: false)

proc scheduleRepeat*(scheduler: Scheduler, interval: Duration, action: TaskAction, initialDelay: Option[Duration] = none(Duration)): TaskHandle =
  if scheduler == nil:
    return nil
  scheduler.nextId.inc
  let startAfter = initialDelay.get(interval)
  let task = ScheduledTask(
    id: scheduler.nextId - 1,
    kind: tkRepeat,
    nextRun: getMonoTime() + startAfter,
    interval: interval,
    action: action,
    steps: @[]
  )
  scheduler.tasks.add(task)
  if scheduler.logger != nil:
    scheduler.logger.debug("Scheduled repeating task", [("id", $task.id), ("interval", $interval.inMilliseconds)])
  TaskHandle(id: task.id, cancelled: false)

proc scheduleSequence*(scheduler: Scheduler, steps: seq[ScheduledStep]): TaskHandle =
  if scheduler == nil or steps.len == 0:
    return nil
  scheduler.nextId.inc
  let firstDelay = steps[0].delay
  let task = ScheduledTask(
    id: scheduler.nextId - 1,
    kind: tkSequence,
    nextRun: getMonoTime() + firstDelay,
    steps: steps,
    stepIndex: 0
  )
  scheduler.tasks.add(task)
  if scheduler.logger != nil:
    scheduler.logger.debug("Scheduled sequence", [("id", $task.id), ("steps", $steps.len)])
  TaskHandle(id: task.id, cancelled: false)

proc tick*(scheduler: Scheduler) =
  ## Evaluate and run due tasks. Intended to be called frequently
  ## from the main message loop.
  if scheduler == nil or scheduler.tasks.len == 0:
    return

  let now = getMonoTime()
  var remaining: seq[ScheduledTask] = @[]

  for task in scheduler.tasks:
    if task.cancelled:
      continue

    if now >= task.nextRun:
      case task.kind
      of tkOnce:
        if task.action != nil:
          task.action()
        task.cancelled = true
      of tkRepeat:
        if task.action != nil:
          task.action()
        task.nextRun = now + task.interval
      of tkSequence:
        if task.stepIndex < task.steps.len:
          let step = task.steps[task.stepIndex]
          if step.action != nil:
            step.action()
          inc task.stepIndex
          if task.stepIndex < task.steps.len:
            task.nextRun = now + task.steps[task.stepIndex].delay
          else:
            task.cancelled = true
    if not task.cancelled:
      remaining.add(task)

  scheduler.tasks = remaining
