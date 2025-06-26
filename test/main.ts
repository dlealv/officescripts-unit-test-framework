
// #region main.ts
// ----------------------------------------
// This script tests the Logging framework implemented in TypeScript for ExcelScript environment.
// It uses the TestRunner class to execute test cases defined in the TestCase class.
// The tests cover various functionalities of the logging framework, including error handling,
// log event creation, and appender functionality.
// 
// The script is designed to run in an Office Scripts/ExcelScript environment, but can also be adapted
// for Node.js or TypeScript environments by removing or modifying certain parts of the code.
// Testing the Logging framework
// ----------------------------------------

/**
 * Entry point for the test suite.
 * @remarks
 * Some tests are adapted for the asynchronous behavior of Office Scripts, requiring a delay to ensure cell values are set before assertions. 
 * The helper `TestCase.runTestAsync(fn)` is used to handle this, detecting the environment via the global variable `RunSyncTest`. 
 * If the variable is set to true, no delay is applied, allowing tests to run synchronously, otherwise a delay is applied to ensure proper execution
 * under Office Scripts. In Node.js/TypeScript environments, this is not required. 
 */
function main(workbook: ExcelScript.Workbook,
) {
  // Parameters and constants definitions
  // ------------------------------------
  const MSG_CELL = "C2" // relative to the active sheet
  //const VERBOSITY = TestRunner.VERBOSITY.OFF        // uncomment the scenario of your preference
  const VERBOSITY = TestRunner.VERBOSITY.HEADER
  //const VERBOSITY = TestRunner.VERBOSITY.SECTION
  const START_TEST = "START TEST" // Used in the title of the test run
  const END_TEST = "END TEST"
  const SHOW_TRACE = false

  let run: TestRunner = new TestRunner(VERBOSITY) // Controles the test execution process specifying the verbosity level
  let success = false // Control variable to send the last message in finally

  // MAIN EXECUTION
  // --------------------

  try {
    // Setting a common layout of the test run proces
    const VERBOSITY_LEVEL = run.getVerbosityLabel()
    run.title(`${START_TEST} with verbosity '${VERBOSITY_LEVEL}'`, 1)
    const INDENT: number = 2 // Use the same indentation level for all test cases

    /*All functions need to be invoked using arrow function (=>).
    Test cases organized by topics. They don't have any dependency, so they can
    be executed in any order.*/

    run.exec("Test Case ScriptError", () => TestCase.utility(), INDENT)
    run.exec("Test Case ScriptError", () => TestCase.scriptError(), INDENT)
    run.exec("Test Case LogEventImpl", () => TestCase.logEventImpl(), INDENT)
    run.exec("Test Case LayoutImpl", () => TestCase.layoutImpl(), INDENT)
    run.exec("Test Case ConsoleAppender", () => TestCase.consoleAppender(), INDENT)
    run.exec("Test Case ExcelAppender", () => TestCase.excelAppender(workbook, MSG_CELL), INDENT)
    // LoggerImpl tests
    run.exec("Test Case LoggerImpl: Counters", () => TestCase.loggerImplCounters(), INDENT)
    run.exec("Test Case LoggerImpl: Export State", () => TestCase.loggerImplExportState(), INDENT)
    run.exec("Test Case LoggerImpl: Reset Singleton", () => TestCase.loggerImplResetSingleton(workbook, MSG_CELL), INDENT)
    run.exec("Test Case LoggerImpl: toString", () => TestCase.loggerImplToString(workbook, MSG_CELL), INDENT)
    run.exec("Test Case LoggerImpl: Lazy Init", () => TestCase.loggerImplLazyInit(), INDENT)
    run.exec("Test Case LoggerImpl: LoggerImpl", () => TestCase.loggerImpl(workbook, MSG_CELL), INDENT)
    run.exec("Test Case LoggerImpl: Internal Errors", () => TestCase.loggerImplScriptError(workbook, MSG_CELL), INDENT)
    success = true
  } catch (e) {
    // TypeScript strict mode: 'e' is of type 'unknown', so we must check its type before property access
    let info: string
    if (e instanceof Error) {
      info = `[${e.name}]: ${e.message}` // Since ScriptError overrided toString method
    } else {
      info = `[unknown]: ${String(e)}`
    }
    success = false
    if (!(e instanceof ScriptError)) { // Unexpected error
      console.log(`Error RAISED`)
      if (SHOW_TRACE) {
        // e is Error here, so stack is safe
        if (e instanceof Error) {
          console.log(`e.stack: ${e.stack}`)
        } else {
          console.log(info)
        }
      } else {
        console.log(info)
      }
    } else { // Handled errors by the framework
      console.log(`ScriptError RAISED`)
      if (SHOW_TRACE) {
        // Safe to call toString if present
        if (typeof e.toString === "function") {
          console.log(`e.toString(): ${e.toString()}`)
        } else {
          console.log(info)
        }
      } else {
        console.log(info)
      }
    }
  } finally {
    TestCase.clear() // safeguard
    run.title(success ? `${END_TEST}: OK` : `${END_TEST}: FAIL`, 1)
  }
} // End of main

// Testing Classes
// -----------------

/**Encapsulates the test cases to be executed as static methods of this class. To be
 * executed via TestRunner.exec method.
 */
class TestCase {

  public static clear(): void { // Clear all the instances
    LoggerImpl.clearInstance()
    ConsoleAppender.clearInstance()
    ExcelAppender.clearInstance()
    AbstractAppender.clearLayout() // Clear the layout
  }

  // Utility to escape regex special characters in variables
  public static escapeRegex(str: string): string {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  }

  /** Removes the timestamp from a string. This is used to compare strings in toString output or to remove
   * the timestamp from the Excel cell value.
   */
  public static removeTimestamp(str: string, token:string=''): string { // Remove timestamp from a string
    let regex = /^(?:\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2},\d{3})\])|"(?:\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2},\d{3})"/g
    return str.replace(regex, token) 
  }

  public static setShortLayout(): void { // Clears and Sets the short layout for the appenders
    AbstractAppender.clearLayout() // You need to set to null first in order to set it via setLayout
    let layout = new LayoutImpl(LayoutImpl.shortFormatterFun)
    AbstractAppender.setLayout(layout) // Set the layout for all appenders
  }

  public static setDefaultLayout(): void { // Clears and Sets the default layout for the appenders
    AbstractAppender.clearLayout() // You need to set to null first in order to set it via setLayout
    let layout = new LayoutImpl() // Default formatter with timestamp
    AbstractAppender.setLayout(layout) // Set the layout for all appenders
  }

  /** Helper method to run a test asynchronously only under Office Scripts/ExcelScript environment.
   * Control via global variable ExcelScriptIsMock. If 'true' it will run the test with no
   * delay, otherwise it will run the test with a delay of 0 milliseconds.
  */
  public static runTestAsync(fn: () => void) {
    // Detect Office Scripts/ExcelScript by presence of ExcelScript global
    if (
      typeof globalThis['ExcelScript'] !== 'undefined' &&
      // @ts-ignore
      globalThis['RunSyncTest'] !== true
    ) {
      setTimeout(fn, 0)
    } else {
      fn()
    }
  }

  /**
   * Returns a new array with only the 'type' and 'message' properties
   * from each LogEvent in the input array.
   * @param logEvents Array of LogEvent objects
   * @returns Array of objects containing only type and message
   */
  public static simplifyLogEvents(logEvents: LogEvent[]): { type: LOG_EVENT; message: string }[] {
    return logEvents.map(event => ({
      type: event.type,
      message: event.message
    }))
  }

  // Helper method to send all possible log event during the testing process consider all possible ACTION value scenrios.
  // It assumes the logger is already initialized. Used in loggerImplLevels.
  public static sendLog(msg: string, type: LOG_EVENT, extraFields: LogEventExtraFields,
    action: typeof LoggerImpl.ACTION[keyof typeof LoggerImpl.ACTION], context: string = "TestCase.sendLog"): void {

    // Defining variables
    let typeStr: string, actionStr: string, errMsg: string, logger: Logger, CONTEXT: string

    logger = LoggerImpl.getInstance() // Get the logger instance (was already instanciated the singleton)
    let level = (logger as LoggerImpl).getLevel() // Get the current level (we can't get it from the input arguments)
    typeStr = LOG_EVENT[type] // Get the level string
    actionStr = LoggerImpl.getActionLabel() // Get the action string
    CONTEXT = `-[type,action]=[${typeStr},${actionStr}]-${context}`
    let extraFieldsStr = extraFields ? ` ${JSON.stringify(extraFields)}` : "" // If extraFields are present, append them to the message
    errMsg = `[${typeStr}] ${msg}${extraFieldsStr}`
    TestCase.runTestAsync(() => {
      Assert.isNotNull(logger, `getInstance()-is not null${CONTEXT}`) // Logger instance should not be null
      if (action === LoggerImpl.ACTION.CONTINUE) { // No ScriptError is thrown, since the action is CONTINUE
        if (type === LOG_EVENT.ERROR) {
          Assert.doesNotThrow(
            () => extraFields ? logger.error(msg, extraFields) : logger.error(msg),
            `error())-do not throw ScriptError${CONTEXT}`
          )
        } else if (type === LOG_EVENT.WARN) {
          Assert.doesNotThrow(
            () => extraFields ? logger.warn(msg, extraFields) : logger.warn(msg),
            `warn()-do not throw ScriptError${CONTEXT}`
          )
        } else if (type === LOG_EVENT.INFO) {
          Assert.doesNotThrow(
            () => extraFields ? logger.info(msg, extraFields) : logger.info(msg),
            `info()-do not throw ScriptError${CONTEXT}`
          )
        } else if (type === LOG_EVENT.TRACE) {
          Assert.doesNotThrow(
            () => extraFields ? logger.trace(msg, extraFields) : logger.trace(msg),
            `trace()-do not throw ScriptError${CONTEXT}`
          )
        } else {// Testing scenario not covered
          throw new AssertionError(`Invalid type: ${typeStr}`)
        }
      } else if (action === LoggerImpl.ACTION.EXIT) { // For error will throw ScriptError, and for warning, depending on the level
        if (type === LOG_EVENT.ERROR) {
          TestCase.setShortLayout()
          Assert.throws(
            () => extraFields ? logger.error(msg, extraFields) : logger.error(msg),
            ScriptError,
            errMsg,
            `error())-throws ScriptError${CONTEXT}`
          )
          TestCase.setDefaultLayout()
        } else if (type === LOG_EVENT.WARN) {
          if (level >= LoggerImpl.LEVEL.WARN) { // If the level is greater than or equal to WARN onl
            TestCase.setShortLayout()
            Assert.throws(
              () => extraFields ? logger.warn(msg, extraFields) : logger.warn(msg),
              ScriptError,
              errMsg,
              `warn()-throws ScriptError${CONTEXT}`
            )
            TestCase.setDefaultLayout()
          } else { // If the level is ERROR then it is not expected to throw ScriptError
            Assert.doesNotThrow(
              () => extraFields ? logger.warn(msg, extraFields) : logger.warn(msg),
              `warn()-do not throw ScriptError${CONTEXT}`
            )
          }
        } else if (type === LOG_EVENT.INFO) {
          Assert.doesNotThrow(
            () => extraFields ? logger.info(msg, extraFields) : logger.info(msg),
            `info()-throws ScriptError${CONTEXT}`
          )
        } else if (type === LOG_EVENT.TRACE) {
          Assert.doesNotThrow(
            () => extraFields ? logger.trace(msg, extraFields) : logger.trace(msg),
            `trace()-throws ScriptError${CONTEXT}`
          )
        } else {
          throw new AssertionError(`Invalid type: ${typeStr}`)
        }
      } else {
        throw new AssertionError(`Invalid action: ${actionStr}`)
      }
    })

  }

  
  /**
   * Helper method to simplify testing scenarios for all possible combinations of LEVEL,ACTION. Except for OFF level.
   */
  private static loggerImplLevels(includeExtraFields: boolean, // If true, it will send extra fields to the log events
    level: typeof LoggerImpl.LEVEL[keyof typeof LoggerImpl.LEVEL],
    action: typeof LoggerImpl.ACTION[keyof typeof LoggerImpl.ACTION],
    workbook: ExcelScript.Workbook, msgCell: string, context: string = "loggerImplLevels"): void {

    // Defining variables
    let logger: Logger, appender: Appender, msgCellRng: ExcelScript.Range,
      activeSheet: ExcelScript.Worksheet, expectedNum: number, actualNum: number,
      levelStr: string, actionStr: string, SUFFIX: string, extraFields: LogEventExtraFields

    // Logger: level, action
    TestCase.clear()
    logger = LoggerImpl.getInstance(level, action)
    levelStr = LoggerImpl.getLevelLabel() // Get the level string
    actionStr = LoggerImpl.getActionLabel() // Get the action string
    SUFFIX = `-[level,action]=[${levelStr},${actionStr}]-${context}, extraFields=${includeExtraFields}` // Suffix for the assertions
    Assert.isNotNull(logger, `LoggerImpl(getInstance)-is not null${SUFFIX}`)

    Assert.equals(logger.getLevel(), level, `getLevel() is correct${SUFFIX}`)
    Assert.equals(logger.getAction(), action, `getAction() is correct${SUFFIX}`)
    // Adding appender
    // Set an Excel appender to the logger
    activeSheet = workbook.getActiveWorksheet()
    msgCellRng = activeSheet.getRange(msgCell)
    appender = ExcelAppender.getInstance(msgCellRng)
    Assert.isNotNull(appender, `ExcelAppender(getInstance) is not null${SUFFIX}`)
    // Adding appender to the logger
    logger!.addAppender(appender)
    Assert.equals(logger!.getAppenders().length, 1, `getAppenders().length-size=1${SUFFIX}`)
    Assert.equals(logger!.getAppenders()[0], appender, `getAppenders()[0]-appender is appender${SUFFIX}`)

    // level is OFF, no log events should be sent to the appender
    if (level === LoggerImpl.LEVEL.OFF) {
      logger.error("This error should not be logged") // log event should not be sent
      logger.warn("This warning should not be logged") // log event should not be sent
      logger.info("This info should not be logged") // log event should not be sent
      logger.trace("This trace should not be logged") // log event should not be sent
      expectedNum = 0 // No log events should be sent to the appender
      actualNum = logger.getCriticalEvents().length
      Assert.equals(actualNum, expectedNum, `getCriticalEvents())--no log events sent${SUFFIX}`)
      Assert.equals(logger.hasErrors(), false, `hasErrors()--no errors logged${SUFFIX}`)
      Assert.equals(logger.hasWarnings(), false, `hasWarnings()--no warnings logged${SUFFIX}`)
      // Checking the last log event sent via AbstractAppender
      Assert.isNull(appender.getLastLogEvent(), `getLastLogEvent()-is null${SUFFIX}`)
      return // No need to continue, since no log events will be sent
    }

    // level is not OFF, so we can continue with the tests
    if (includeExtraFields) {
      extraFields = { userId: 123, sessionId: "abc" }
    } else {
      extraFields = undefined
    }

    repeatCheckPerLevel(level, `repeatCheckPerLevel-extraFields=${includeExtraFields}`)
    // Inner functions
    function repeatCheckingCriticalEvents(msg: string, type: LOG_EVENT, context: string = "repeatCheckingCriticalEvents"): void {
      let CONTEXT = `-[level,action]=[${levelStr},${actionStr}]-${context}`
      // Checking the number of critical events sent
      // Checking the critical event sent
      Assert.isNotNull(logger.getCriticalEvents(), `getCriticalEvents()-critial events are not null${CONTEXT}`)
      let lastIdx = logger.getCriticalEvents().length - 1
      let actualEvent = logger.getCriticalEvents()[lastIdx]
      TestCase.runTestAsync(() => {
        Assert.isTrue(lastIdx >= 0, `LoggerImpl(getCriticalEvents)-critical events array is not empty${CONTEXT}`)
        Assert.isNotNull(actualEvent, `getCriticalEvents()-last log event is not null${CONTEXT}`)
        Assert.equals(actualEvent!.type, type, `getCriticalEvents()-last log event type is correct${CONTEXT}`)
        Assert.equals(actualEvent!.message, msg, `getCriticalEvents()-last log event message is correct${CONTEXT}`)
      })
    }

    function repeatCheckingAbstractAppender(expectedMsg: string, expectedType: LOG_EVENT, context: string = "repeatCheckingAbstractAppender"): void {
      let CONTEXT = `-[level,action]=[${levelStr},${actionStr}]-${context}`
      // Checking the last log event sent via AbstractAppender
      TestCase.runTestAsync(() => {
        Assert.isNotNull(appender.getLastLogEvent(), `getLastLogEvent()-is not null${CONTEXT}`)
        let actualEvent = appender.getLastLogEvent()
        Assert.isNotNull(actualEvent, `getLastLogEvent()-last log event is not null${CONTEXT}`)
        Assert.equals(actualEvent!.type, expectedType, `getLastLogEvent()-last log event type is ir correct${CONTEXT}`)
        Assert.equals(actualEvent!.message, expectedMsg, `getLastLogEvent()-last log event message is correct${CONTEXT}`)
      })
    }

    function repeatCheckingExcelCellContent(expectedMsg: string, expectedType: LOG_EVENT, context: string = "repeatCheckingExcelCellContent"): void {
      // Checking the content of the excel cell
      let CONTEXT = `-[level,action]=[${levelStr},${actionStr}]-${context}`
      const TK = "<TIMESTAMP>" // Token to be used to remove the timestamp from the message
      Assert.isNotNull(msgCellRng, `ExcelAppender(getInstance)-msgCellRng is not null${CONTEXT}`)
      let value = msgCellRng.getValue().toString()
      let actualMsg = TestCase.removeTimestamp(value, TK) // under default configuration the output has stimestamp
      const PREFIX = actualMsg == value ? "" : `${TK} ` // Adjust it in case timestamp was removed
      let extraFieldsStr = extraFields ? ` ${JSON.stringify(extraFields)}` : "" // If extraFields are present, append them to the message
      expectedMsg = `${PREFIX}[${LOG_EVENT[expectedType]}] ${expectedMsg}${extraFieldsStr}`
      TestCase.runTestAsync(() => {
        Assert.equals(actualMsg, expectedMsg, `ExcelAppender(msgCellRng.getValue)-excel cell content is correct${CONTEXT}`)
      })
    }

    function repeatCheckPerLevel(level: typeof LoggerImpl.LEVEL[keyof typeof LoggerImpl.LEVEL], context: string = "repeatCheckPerLevel"): void {
      // Defining variables
      let addCriticalEvent: boolean, addEvent: boolean, lastCriticalMsg: string, lastCriticalType: LOG_EVENT,
        lastType: LOG_EVENT, expectedMsg: string, expectedType: LOG_EVENT, lastMsg: string

      let CONTEXT = `[level,action]=[${levelStr},${actionStr}]-${context}` // to be used for calling other inner functions
      let SUFFIX = `-${CONTEXT}` // Used for final assertions

      // Sending error log event
      lastMsg = `Error message(${levelStr},${actionStr})`
      lastType = LOG_EVENT.ERROR
      TestCase.sendLog(lastMsg, lastType, extraFields, action, CONTEXT) // Depending on action, could throw ScriptError or not
      expectedNum = 1 // error log event is always a critical event
      actualNum = logger.getCriticalEvents().length
      TestCase.runTestAsync(() => {
        Assert.equals(actualNum, expectedNum, `getCriticalEvents(length=1)-from error event is correct${SUFFIX}`)
        Assert.equals(logger.hasErrors(), true, `hasErrors() is true${SUFFIX}`)
        Assert.equals(logger.hasWarnings(), false, `hasWarnings() is false${SUFFIX}`)
      })
      repeatCheckingCriticalEvents(lastMsg, lastType)
      repeatCheckingAbstractAppender(lastMsg, lastType, CONTEXT)
      repeatCheckingExcelCellContent(lastMsg, lastType, CONTEXT)
      lastCriticalMsg = lastMsg
      lastCriticalType = lastType

      // Sending warning log event
      expectedMsg = `Warning event(${levelStr},${actionStr})`
      expectedType = LOG_EVENT.WARN
      TestCase.sendLog(expectedMsg, expectedType, extraFields, action, CONTEXT)
      expectedNum = level > LoggerImpl.LEVEL.ERROR ? 2 : 1 // If level is ERROR, only the error log event was sent
      actualNum = logger.getCriticalEvents().length
      TestCase.runTestAsync(() => {
        Assert.equals(actualNum, expectedNum, `getCriticalEvents(length) from warning event is correct${SUFFIX}`)
      })

      addCriticalEvent = level >= LoggerImpl.LEVEL.WARN ? true : false // If level is WARN, warning log event was sent
      addEvent = level >= LoggerImpl.LEVEL.WARN ? true : false // If level is WARN or lower, warning log event was sent
      TestCase.runTestAsync(() => {
        Assert.isTrue(logger.hasErrors(), `hasErrors() is true${SUFFIX}`)
        Assert.equals(logger.hasWarnings(), addEvent, `hasWarnings() is correct${SUFFIX}`)
      })

      if (addCriticalEvent) {
        lastCriticalMsg = expectedMsg
        lastCriticalType = expectedType
      }
      if (addEvent) {
        lastMsg = expectedMsg
        lastType = expectedType
      }
      repeatCheckingCriticalEvents(lastCriticalMsg, lastCriticalType, CONTEXT)
      repeatCheckingAbstractAppender(lastMsg, lastType, CONTEXT)
      repeatCheckingExcelCellContent(lastMsg, lastType, CONTEXT)

      // Sending info log event
      expectedMsg = `Info event(${level},${action})`
      expectedType = LOG_EVENT.INFO
      TestCase.sendLog(expectedMsg, expectedType, extraFields, action, CONTEXT)
      actualNum = logger.getCriticalEvents().length
      TestCase.runTestAsync(() => {
        Assert.equals(actualNum, expectedNum, `getCriticalEvents(length)-from info log is correct${SUFFIX}`)
        Assert.equals(logger.hasErrors(), true, `hasErrors() is true${SUFFIX}`)
        Assert.equals(logger.hasWarnings(), addCriticalEvent, `hasWarnings() is correct${SUFFIX}`)
      })
      addEvent = level >= LoggerImpl.LEVEL.INFO ? true : false // If level is INFO or lower, info log event was sent
      if (addEvent) {
        lastMsg = expectedMsg
        lastType = expectedType
      }
      repeatCheckingCriticalEvents(lastCriticalMsg, lastCriticalType, CONTEXT)    // warning event is the last one
      repeatCheckingAbstractAppender(lastMsg, lastType, CONTEXT)
      repeatCheckingExcelCellContent(lastMsg, lastType, CONTEXT)

      // Sending trace log event
      expectedMsg = `Trace event(${level},${action})`
      expectedType = LOG_EVENT.TRACE
      TestCase.sendLog(expectedMsg, expectedType, extraFields, action, CONTEXT)
      actualNum = logger.getCriticalEvents().length
      TestCase.runTestAsync(() => {
        Assert.equals(actualNum, expectedNum, `getCriticalEvents(length) from trace event is correct${SUFFIX}`)
        Assert.equals(logger.hasErrors(), true, `hasErrors() is true${SUFFIX}`)
        Assert.equals(logger.hasWarnings(), addCriticalEvent, `hasWarnings() is correct${SUFFIX}`)
      })
      addEvent = level >= LoggerImpl.LEVEL.TRACE ? true : false // If level is TRACE or lower, trace log event was sent
      if (addEvent) {
        lastMsg = expectedMsg
        lastType = expectedType
      }
      repeatCheckingCriticalEvents(lastCriticalMsg, lastCriticalType, CONTEXT)    // warning event is the last one
      repeatCheckingAbstractAppender(lastMsg, lastType, CONTEXT)
      repeatCheckingExcelCellContent(lastMsg, lastType, CONTEXT)

    }
  }

  // TEST CASES

  public static utility(): void { // Unit tests for utility methods

    // Defining the variables to be used in the tests
    let expectedStr: string, actualStr: string, msg: string, errMsg: string

    // Testing data2Str
    let date = new Date(2025, 0, 1, 1, 1, 1, 1) // January is 0 in JavaScript Date
    actualStr = Utility.date2Str(date)
    expectedStr = `2025-01-01 01:01:01,001`
    Assert.equals(actualStr, expectedStr, "utility(data2Str)")

    // Testing data2Str with null
    actualStr = Utility.date2Str(null as unknown as Date)
    expectedStr = `[Utility.date2Str]: Invalid Date`
    Assert.equals(actualStr, expectedStr, "utility(data2Str) - null date")

    // Testing data2Str with undefined
    actualStr = Utility.date2Str(undefined as unknown as Date)
    expectedStr = `[Utility.date2Str]: Invalid Date`
    Assert.equals(actualStr, expectedStr, "utility(data2Str) - undefined date")

    // Testing data2Str with invalid date
    actualStr = Utility.date2Str(new Date("invalid date"))
    expectedStr = `[Utility.date2Str]: Invalid Date`
    Assert.equals(actualStr, expectedStr, "utility(data2Str) - invalid date")


    // Testing validateLogEventFactory
    const validFactory: LogEventFactory = (message: string, eventType: LOG_EVENT) => {
      return new LogEventImpl(message, eventType)
    }
    Assert.doesNotThrow(
      () => Utility.validateLogEventFactory(validFactory),
      "utility(validateLogEventFactory) - valid factory"
    )

    // Testing validateLogEventFactory with null
    errMsg = "Invalid <anonymous>: Not a function"
    Assert.throws(
      () => Utility.validateLogEventFactory(null as unknown as LogEventFactory),
      ScriptError,
      errMsg,
      "utility(validateLogEventFactory) - null factory"
    )

    // Testing validateLogEventFactory with undefined
    Assert.throws(
      () => Utility.validateLogEventFactory(undefined as unknown as LogEventFactory),
      ScriptError,
      errMsg,
      "utility(validateLogEventFactory) - undefined factory"
    )

    // Testing validateLogEventFactory with non-function
    Assert.throws(
      () => Utility.validateLogEventFactory("invalid" as unknown as LogEventFactory),
      ScriptError,
      errMsg,
      "utility(validateLogEventFactory) - non-function-string"
    )

    // Testing validateLogEventFactory with non-function
    Assert.throws(
      () => Utility.validateLogEventFactory(123 as unknown as LogEventFactory),
      ScriptError,
      errMsg,
      "utility(validateLogEventFactory) - non-function-number"
    )

    // Testing validateLogEventFactory with non-function
    Assert.throws(
      () => Utility.validateLogEventFactory({} as unknown as LogEventFactory),
      ScriptError,
      errMsg,
      "utility(validateLogEventFactory) - non-function-object"
    )

    // Testing validateLogEventFactory with non-function
    Assert.throws(
      () => Utility.validateLogEventFactory([] as unknown as LogEventFactory),
      ScriptError,
      errMsg,
      "utility(validateLogEventFactory) - non-function-array"
    )

    // Testing providing the funName and context for a non valid function
    errMsg = "[TestCase.utility]: Invalid non-function-string: Not a function"
    Assert.throws(
      () => Utility.validateLogEventFactory("invalid" as unknown as LogEventFactory, "non-function-string", "TestCase.utility"),
      ScriptError,
      errMsg,
      "utility(validateLogEventFactory) - non-function-string"
    )
    // Note: There is no way to ckeck the arity of a function in JavaScript, so we can't test it here.

  }

  public static scriptError(): void { // Unit tests for the ScriptError class
    TestCase.clear() // Clear all the instances

    // Defining the variables to be used in the tests
    let expectedMsg: string, actualMsg: string, cause: Error, origin: ScriptError

    // Testing raising a ScriptError without cause
    expectedMsg = "Script Error message"
    Assert.throws(
      () => { throw new ScriptError(expectedMsg) },
      ScriptError,
      expectedMsg,
      "scriptError(notcause)"
    )

    // Testing raising a ScriptError with cause
    cause = new TypeError("Type Error message")
    origin = new ScriptError(expectedMsg, cause)
    expectedMsg = "Script Error message (caused by 'TypeError' with message 'Type Error message')"
    Assert.throws(
      () => { throw origin },
      ScriptError,
      expectedMsg,
      "scriptError(with cause)"
    )

    function buildRegex(trigger: ScriptError): RegExp {// Building regex for toString
      let NAME = trigger.cause ? trigger.cause.name : trigger.name
      let MSG = trigger.cause ? trigger.cause.message : trigger.message
      const regex = new RegExp(
        `^${TestCase.escapeRegex(trigger.name)}: ${TestCase.escapeRegex(trigger.message)}\\n` +  // Header
        `Stack trace:\\n` +                                                    // Stack section
        `${TestCase.escapeRegex(NAME)}: ${TestCase.escapeRegex(MSG)}\\n` +                       // type and message
        `( +at .+\\n?)+$`                                                      // Variable stack trace lines
      )
      return regex
    }

    let scriptError = new ScriptError("Script Error message")
    let scriptErrorwithCause = new ScriptError("Script Error message", cause)

    // Testing without cause
    let regex = buildRegex(scriptError)
    Assert.equals(regex.test(scriptError.toString()), true,
      "scriptError(toString without cause)"
    )
    // Testing with cause
    regex = buildRegex(scriptErrorwithCause)
    Assert.equals(regex.test(scriptErrorwithCause.toString()), true,
      "scriptError(toString with cause)"
    )

    // Testing rethrowCauseIfNeeded
    // Testing rethrowCauseIfNeeded without cause
    expectedMsg = "Script Error message"
    try {
      const err = new ScriptError(expectedMsg)
      err.rethrowCauseIfNeeded()
      Assert.fail("Expected ScriptError to be thrown")
    } catch (e) {
      isScriptError(e, "LogEvent(rethrowCauseIfNeeded)-No cause")
      Assert.equals((e as ScriptError).message, expectedMsg, "LogEvent(rethrowCauseIfNeeded)-Top-level error")
    }

    // Cause is not a ScriptError, so it should be rethrown
    try {
      cause = new Error("Root cause")
      origin = new ScriptError("Wrapper error", cause)
      origin.rethrowCauseIfNeeded()
      //Assert.fail("Expected root cause Error to be thrown")
    } catch (e) {
      isError(e, "LogEvent(rethrowCauseIfNeeded)-Root cause is Error")
      isNotScriptError(e, "LogEvent(rethrowCauseIfNeeded)-Root cause is not a ScriptError")
      Assert.equals((e as Error).message, "Root cause")
    }

    // Deepest cause is a ScriptError, so it should be rethrown
    try {
      const root = new Error("Root error")
      const nested = new ScriptError("Nested script error", root)
      const top = new ScriptError("Top script error", nested)
      top.rethrowCauseIfNeeded()
      Assert.fail("Expected root Error to be thrown")
    } catch (e) {
      isError(e, "LogEvent(rethrowCauseIfNeeded)-Deepest cause is Error")
      isNotScriptError(e, "LogEvent(rethrowCauseIfNeeded)-Deepest cause is not a ScriptError")
      Assert.equals((e as Error).message, "Root error")
    }

    TestCase.clear() // Clear all the instances

    // Inner functions
    function isScriptError(value: unknown, message?: string): asserts value is ScriptError {
      if (!(value instanceof ScriptError)) {
        const prefix = message ? `${message}: ` : ""
        throw new AssertionError(
          `${prefix}Expected instance of ScriptError, but got ${(value as object)?.constructor?.name || typeof value}`
        )
      }
    }

    function isNotScriptError(value: unknown, message?: string): void {
      if (value instanceof ScriptError) {
        const prefix = message ? `${message}: ` : ""
        throw new AssertionError(
          `${prefix}Expected value NOT to be instance of ScriptError, but got ${(value as object)?.constructor?.name || typeof value}`
        )
      }
    }

    function isError(value: unknown, message?: string): asserts value is Error {
      if (!(value instanceof Error)) {
        const prefix = message ? `${message}: ` : ""
        throw new AssertionError(
          `${prefix}Expected instance of Error, but got ${(value as object)?.constructor?.name || typeof value}`
        )
      }
    }

    function isNotError(value: unknown, message?: string): void {
      if (value instanceof Error) {
        const prefix = message ? `${message}: ` : ""
        throw new AssertionError(
          `${prefix}Expected value NOT to be instance of Error, but got ${(value as object)?.constructor?.name || typeof value}`
        )
      }
    }

  }

  public static layoutImpl(): void { // Unit tests for LayoutImpl class
    TestCase.clear()

    // Deffining the variables to be used in the tests
    let layout: Layout, event: LogEvent, actualStr: string, expectedStr: string, expectedMsg: string, expectedType: LOG_EVENT, eventWithExtras: LogEvent,
    errMsg: string

    expectedMsg = "Test message"
    expectedType = LOG_EVENT.INFO
    event = new LogEventImpl(expectedMsg, expectedType)
    const TK= "<TIMESTAMP>" // Token to be used to replace the timestamp from the message

    // Testing constructor: short layout
    layout = new LayoutImpl(LayoutImpl.shortFormatterFun) // with short formatter
    Assert.isNotNull(layout, "LayoutImpl(constructor-short layout is not null)")
    Assert.equals((layout as LayoutImpl).getFormatter(), LayoutImpl.shortFormatterFun, "LayoutImpl(constructor-getFormatter() short formatter)")

    // Testing constructor: long layout
    layout = new LayoutImpl() // Default formatter with timestamp
    Assert.isNotNull(layout, "LayoutImpl(constructor-long layout is not null)")
    Assert.equals((layout as LayoutImpl).getFormatter(), LayoutImpl.defaultFormatterFun, "LayoutImpl(constructor-getFormatter() long formatter)")
    // Testing constructor: invalid formatter, since the input argument was provided, it doesn't use the default formatter
    expectedStr = `[LayoutImpl.constructor]: Invalid Layout: The internal '_formatter' ` +
      `property must be a function accepting a single LogEvent argument. See LayoutImpl documentation for usage.`
    Assert.throws(
      () => new LayoutImpl("Invalid formatter" as unknown as (event: LogEvent) => string),
      ScriptError,
      expectedStr,
      "LayoutImpl(ScriptError)-constructor - invalid formatter"
    )

    //Testing constructor: null formatter:  null is valid, since it defaults to default formatter
    Assert.doesNotThrow(() => {
      new LayoutImpl(null as unknown as (event: LogEvent) => string)
    },
      "LayoutImpl(ScriptError)-constructor - null formatter")

    //Testing constructor: undefined formatter:  undefined is valid, since it defaults to default formatter
    Assert.doesNotThrow(() => {
      new LayoutImpl(null as unknown as (event: LogEvent) => string)
    },
      "LayoutImpl(ScriptError)-constructor - undefined formatter")

      // Testing constructor: invalid formatter, return empty string
    errMsg = "[LayoutImpl.constructor]: Formatter function must return a non-empty string for a valid LogEvent. Got: empty string"
    Assert.throws(
      () => new LayoutImpl(function alwaysEmpty(e) { return "" }),
      ScriptError,
      errMsg,
      "LayoutImpl(constructor) - formatter returns empty string"
    )

    // Testing constructor: invalid formatter return null
    errMsg = "[LayoutImpl.constructor]: Formatter function must return a non-empty string for a valid LogEvent. Got: null"
    Assert.throws(
      () => new LayoutImpl(function alwaysNull(e) { return null as unknown as string }),
      ScriptError,
      errMsg,
      "LayoutImpl(constructor) - formatter returns null"
    )

    //Testing constructor: formatter with wrong arity (no arguments)
    errMsg = `[LayoutImpl.constructor]: Invalid Layout: The internal '_formatter' property must be a function accepting a single LogEvent argument. See LayoutImpl documentation for usage.`
    Assert.throws(
      () => new LayoutImpl(function zeroArgs() { return "ok" }),
      ScriptError,
      errMsg,
      "LayoutImpl(constructor) - formatter with zero args"
    )

    //Testing constructor: formatter with wrong arity (2+ arguments)
    errMsg = `[LayoutImpl.constructor]: Invalid Layout: The internal '_formatter' property must be a function accepting a single LogEvent argument. See LayoutImpl documentation for usage.`
    Assert.throws(
      () => new LayoutImpl(function twoArgs(e:number, f:string) { return "ok" } as unknown as LayoutFormatter),
      ScriptError,
      errMsg,
      "LayoutImpl(constructor) - formatter with two args"
    )


    // Testing format with short formatter
    layout = new LayoutImpl(LayoutImpl.shortFormatterFun) // with short formatter
    expectedStr = `[${LOG_EVENT[expectedType]}] ${expectedMsg}`
    actualStr = TestCase.removeTimestamp(layout.format(event))
    Assert.equals(actualStr, expectedStr, "LayoutImpl(format-short formatter)")

    // Testing format with short formatter and with extra fields
    eventWithExtras = new LogEventImpl(expectedMsg, expectedType, { userId: 123, sessionId: "abc" })
    expectedStr = `[${LOG_EVENT[expectedType]}] ${expectedMsg} {"userId":123,"sessionId":"abc"}`
    actualStr = layout.format(eventWithExtras)
    Assert.equals(actualStr, expectedStr, "LayoutImpl(format-short formatter with extras)")

    // Testing format with long formatter
    expectedStr = `${TK} [${LOG_EVENT[expectedType]}] ${expectedMsg}`
    layout = new LayoutImpl() // Default formatter with timestamp
    actualStr = TestCase.removeTimestamp(layout.format(event), TK)
    Assert.equals(actualStr, expectedStr, "LayoutImpl(format-long formatter)")

    // Testing format with long formatter and with extra fields
    eventWithExtras = new LogEventImpl(expectedMsg, expectedType, { userId: 123, sessionId: "abc" })
    expectedStr = `${TK} [${LOG_EVENT[expectedType]}] ${expectedMsg} {"userId":123,"sessionId":"abc"}`
    actualStr = TestCase.removeTimestamp(layout.format(eventWithExtras), TK)
    Assert.equals(actualStr, expectedStr, "LayoutImpl(format-long formatter with extras)")

    // Testing format with invalid LogEvent
    layout = new LayoutImpl() // or pick any valid formatter
    errMsg = "[LayoutImpl.format]: LogEvent.type='null' property must be a number (LOG_EVENT enum value)."
    Assert.throws(
      () => layout.format({ type: null, message: "msg", timestamp: new Date(), extraFields: {} } as unknown as LogEvent),
      ScriptError,
      errMsg,
      "LayoutImpl(format) - invalid LogEvent (type=null)"
    )

    // Testing toString with short formatter
    layout = new LayoutImpl(LayoutImpl.shortFormatterFun) // with short formatter
    expectedStr = `LayoutImpl: {formatter: [Function: "shortLayoutFormatterFun"]}`
    actualStr = layout.toString()
    Assert.equals(actualStr, expectedStr, "LayoutImpl(toString-short formatter)")

    // Testing toString with long formatter
    layout = new LayoutImpl() // Default formatter with timestamp
    expectedStr = `LayoutImpl: {formatter: [Function: "defaultLayoutFormatterFun"]}`
    actualStr = layout.toString()
    Assert.equals(actualStr, expectedStr, "LayoutImpl(toString-long formatter)")

    // Testing toString with a custom formatter
    layout = new LayoutImpl(function customNamedFormatter(e: LogEvent) { return "X" })
    Assert.equals(layout.toString(), `LayoutImpl: {formatter: [Function: "customNamedFormatter"]}`, "LayoutImpl(toString) - custom named formatter")

    // Testing validateLayout: invalid formatter: null
    expectedStr = `[LayoutImpl.constructor]: Invalid Layout: layout object is null or undefined`
    Assert.throws(
      () => LayoutImpl.validateLayout(null, "LayoutImpl.constructor"),
      ScriptError,
      expectedStr,
      "LayoutImpl(validateLayout)-null layout"
    )

    // Testing validateLayout: invalid formatter: undefined
    expectedStr = `[LayoutImpl.constructor]: Invalid Layout: layout object is null or undefined`
    Assert.throws(
      () => LayoutImpl.validateLayout(undefined, "LayoutImpl.constructor"),
      ScriptError,
      expectedStr,
      "LayoutImpl(validateLayout)-undefined layout"
    )

    // Testing validateLayout: invalid formatter: string is not a function
    expectedStr = `[LayoutImpl.constructor]: Invalid Layout: The internal '_formatter' ` +
      `property must be a function accepting a single LogEvent argument. See LayoutImpl documentation for usage.`
    const customInvalidFormatter = "Invalid formatter" as unknown as (event: LogEvent) => string
    Assert.throws(
      () => LayoutImpl.validateLayout(new LayoutImpl(customInvalidFormatter), "TestCase.layoutImpl"),
      ScriptError,
      expectedStr,
      "LayoutImpl(validateLayout)-validateLayout - string is not a function"
    )

    // Testing validateLayout: valid layout
    const customValidFormatter = (event: LogEvent) => `${event.type}: ${event.message}`
    Assert.doesNotThrow(
      () => LayoutImpl.validateLayout(new LayoutImpl(customValidFormatter), "LayoutImpl.validateLayout"),
      "LayoutImpl(validateLayout)-valid layout with custom formatter"
    )

    TestCase.clear()

  }

  public static logEventImpl(): void { // Unit tests for LogEventImpl class
    TestCase.clear()

    // Defining the variables to be used in the tests
    let actualEvent: LogEvent, expectedMsg: string, actualMsg: string, expectedStr: string, actualStr: string,
      expectedType: LOG_EVENT, actualType: LOG_EVENT, actualtimeStamp: Date, errMsg: string,
      expectedTimestamp: Date

    let eventExtras: LogEventExtraFields = { userId: 123, sessionId: "abc" }

    // Testing constructor
    expectedMsg = "Test message testing logEventImpl"
    expectedType = LOG_EVENT.INFO
    expectedTimestamp = new Date()
    actualEvent = new LogEventImpl(expectedMsg, expectedType)
    Assert.isNotNull(actualEvent, "LogEventImpl(constructor-is not null)")

    // Testing constructor with extra fields
    let eventWithExtras = new LogEventImpl(expectedMsg, expectedType, eventExtras, expectedTimestamp)
    Assert.isNotNull(eventWithExtras, "LogEventImpl(constructor with extras)-is not null")
    Assert.equals(eventWithExtras.message, expectedMsg, "LogEventImpl(constructor with extras)-message is correct")
    Assert.equals(eventWithExtras.type, expectedType, "LogEventImpl(constructor with extras)-type is correct")
    Assert.isNotNull(eventWithExtras.timestamp, "LogEventImpl(constructor with extras)-timestamp is not null")
    Assert.equals(eventWithExtras.extraFields.userId, eventExtras.userId, "LogEventImpl(constructor with extras)-userId is correct")
    Assert.equals(eventWithExtras.extraFields.sessionId, eventExtras.sessionId, "LogEventImpl(constructor with extras)-sessionId is correct")

    // Testing the constructorw with no extra field and checking the value of the property
    actualEvent = new LogEventImpl(expectedMsg, expectedType)
    Assert.isNotNull(actualEvent.extraFields, "LogEventImpl(constructor with no extras)-extraFields is not null")
    Assert.equals(Object.keys(actualEvent.extraFields).length, 0, "LogEventImpl(constructor with no extras)-extraFields is empty")

    // Testing extraFields with a field that is a function
    eventExtras = { userId: 123, sessionId: "abc", logTime: () => new Date().toISOString() }
    eventWithExtras = new LogEventImpl(expectedMsg, expectedType, eventExtras, expectedTimestamp)
    Assert.isNotNull(eventWithExtras, "LogEventImpl(constructor with function extra)-is not null")
    Assert.equals(eventWithExtras.message, expectedMsg, "LogEventImpl(constructor with function extra)-message is correct")
    Assert.equals(eventWithExtras.type, expectedType, "LogEventImpl(constructor with function extra)-type is correct")
    Assert.isNotNull(eventWithExtras.timestamp, "LogEventImpl(constructor with function extra)-timestamp is not null")
    Assert.equals(eventWithExtras.extraFields.userId, eventExtras.userId, "LogEventImpl(constructor with function extra)-userId is correct")
    Assert.equals(eventWithExtras.extraFields.sessionId, eventExtras.sessionId, "LogEventImpl(constructor with function extra)-sessionId is correct")

    // Testing constructor as undefined
    Assert.doesNotThrow(
      () => new LogEventImpl(expectedMsg, LOG_EVENT.INFO, undefined as unknown as LogEventExtraFields, new Date()),
      "LogEventImpl(ScriptError)-constructor - undefined extraFields"
    )

    // Testing properties of the LogEventImpl created
    actualMsg = (actualEvent as LogEvent).message
    actualType = (actualEvent as LogEvent).type
    actualtimeStamp = (actualEvent as LogEvent).timestamp
    Assert.isNotNull(actualtimeStamp, "LogEventImpl(get timestamp) is not null")
    Assert.equals(actualType, expectedType, "LogEventImpl(get type())")
    Assert.equals(actualMsg, expectedMsg, "LogEventImpl(get message())")

    // Teesting constructor with invalid event type
    errMsg = "[LogEventImpl.constructor]: LogEvent.type='-1' property is not defined in the LOG_EVENT enum."
    Assert.throws(
      () => new LogEventImpl(expectedMsg, -1 as LOG_EVENT),
      ScriptError,
      errMsg,
      "LogEventImpl(ScriptError)-constructor - non valid LOG_EVENT"
    )

    // Testing constructor with null message
    expectedMsg = null as unknown as string // null is not a valid message
    errMsg = "[LogEventImpl.constructor]: LogEvent.message='null' property must be a string."
    Assert.throws(
      () => new LogEventImpl(expectedMsg, LOG_EVENT.INFO),
      ScriptError,
      errMsg,
      "LogEventImpl(ScriptError)-constructor - null message"
    )

    // Testing constructor with undefined message
    expectedMsg = undefined as unknown as string // undefined is not a valid message
    errMsg = "[LogEventImpl.constructor]: LogEvent.message='undefined' property must be a string."
    Assert.throws(
      () => new LogEventImpl(expectedMsg, LOG_EVENT.INFO),
      ScriptError,
      errMsg,
      "LogEventImpl(ScriptError)-constructor - undefined message"
    )

    // Testing constructor with an empty string
    expectedMsg = "" // empty string is not a valid message
    errMsg = "[LogEventImpl.constructor]: LogEvent.message cannot be empty."
    Assert.throws(
      () => new LogEventImpl(expectedMsg, LOG_EVENT.INFO),
      ScriptError,
      errMsg,
      "LogEventImpl(ScriptError)-constructor - empty message"
    )

    // Testing constructor with white space only message
    Assert.throws(
      () => new LogEventImpl("   ", LOG_EVENT.INFO, {}, new Date()),
      ScriptError,
     errMsg,
      "LogEventImpl(constructor) - whitespace message throws"
    )

    // Testing Constructor with non valid date
    errMsg = "[LogEventImpl.constructor]: LogEvent.timestamp='null' property must be a Date."
    expectedMsg = "Test message"
    Assert.throws(
      () => new LogEventImpl(expectedMsg, LOG_EVENT.INFO, {}, null as unknown as Date),
      ScriptError,
      errMsg,
      "LogEventImpl(ScriptError)-constructor - null timestamp"
    )

    // Testing constructor with wrong extraFields: null
    errMsg = "[LogEventImpl.constructor]: extraFields must be a plain object."
    Assert.throws(
      () => new LogEventImpl(expectedMsg, LOG_EVENT.INFO, null as unknown as LogEventExtraFields, new Date()),
      ScriptError,
      errMsg,
      "LogEventImpl(ScriptError)-constructor - null extraFields"
    )

    // Testing constructor with wrong extraFields: non valid function
    errMsg = "[LogEventImpl.constructor]: extraFields must be a plain object."
    Assert.throws(
      () => new LogEventImpl(expectedMsg, LOG_EVENT.INFO, "invalid" as unknown as LogEventExtraFields, new Date()),
      ScriptError,
      errMsg,
      "LogEventImpl(ScriptError)-constructor - non valid extraFields"
    )

    // Testing constructor for extraFields with Deep object value:
    // Deep object value
    errMsg = "[LogEventImpl.constructor]: extraFields[bad] has invalid type: object. Must be string, number, Date, or function."
    Assert.throws(
      () => new LogEventImpl("msg", LOG_EVENT.INFO, { bad: { nested: true } } as unknown as LogEventExtraFields, new Date()),
      ScriptError,
      errMsg,
      "LogEventImpl(constructor) - extraFields with deep object throws"
    )

    // Testing constructor for extraFields with array value:
    errMsg = "[LogEventImpl.constructor]: extraFields[bad] has invalid type: object. Must be string, number, Date, or function."
    Assert.throws(
      () => new LogEventImpl("msg", LOG_EVENT.INFO, { bad: [1, 2, 3] } as unknown as LogEventExtraFields, new Date()),
      ScriptError,
      errMsg,
      "LogEventImpl(constructor) - extraFields with array throws"
    )

    // Testing constructor with extraFields with an undefined value
    errMsg = "[LogEventImpl.constructor]: extraFields[bad] must not be undefined."
    Assert.throws(
      () => new LogEventImpl("msg", LOG_EVENT.INFO, { bad: undefined }, new Date()),
      ScriptError,
      errMsg,
      "LogEventImpl(constructor) - extraFields with undefined throws"
    )

    // Testing toString
    let TK = "<TIMESTAMP>"
    let regex: RegExp = new RegExp(`^LogEventImpl: {timestamp=${TK}, type="${LOG_EVENT[actualType]}", message="${actualMsg}"}$`)
    expectedStr = `[${actualType}] ${expectedMsg}`
    actualStr = TestCase.removeTimestamp((actualEvent as LogEvent).toString(), TK) // Remove timestamp for comparison
    Assert.equals(regex.test(actualStr), true, "LogEventImpl(toString())")

    // Testing toString with extra fields
    expectedStr = `LogEventImpl: {timestamp="${Utility.date2Str(expectedTimestamp)}", type="${LOG_EVENT[actualType]}", message="${actualMsg}", extraFields=${JSON.stringify(eventExtras)}}`
    actualStr = (eventWithExtras as LogEvent).toString()
    Assert.equals(actualStr, expectedStr, "LogEventImpl(toString with extras)")

    // Testing eventToLabel, valid case
    // It doesn't
    expectedStr = `INFO`
    actualStr = LogEventImpl.eventTypeToLabel(actualType)
    Assert.equals(actualStr, expectedStr, "LogEventImpl(eventTypeToLabel)")


    // Testing validateLogEvent (exception cases): null
    errMsg = "[LogEventImpl.validateLogEvent]: LogEvent.type='null' property must be a number (LOG_EVENT enum value)."
    Assert.throws(
      () => LogEventImpl.validateLogEvent({ type: null, message: actualMsg, timestamp: actualtimeStamp }),
      ScriptError,
      errMsg,
      "LogEventImpl(validateLogEvent)-null type"
    )
    // Testing validateLogEvent (exception cases): undefined
    errMsg = "[LogEventImpl.validateLogEvent]: LogEvent.type='undefined' property must be a number (LOG_EVENT enum value)."
    Assert.throws(
      () => LogEventImpl.validateLogEvent({ type: undefined, message: actualMsg, timestamp: actualtimeStamp }),
      ScriptError,
      errMsg,
      "LogEventImpl(validateLogEvent)-undefined type"
    )

    // Testing validateLogEvent (exception cases): null message
    errMsg = "[LogEventImpl.validateLogEvent]: LogEvent.message='null' property must be a string."
    Assert.throws(
      () => LogEventImpl.validateLogEvent({ type: actualType, message: null as unknown as string, timestamp: actualtimeStamp }),
      ScriptError,
      errMsg,
      "LogEventImpl(validateLogEvent)-null message"
    )

    // Testing validateLogEvent (exception cases): undefined message
    errMsg = "[LogEventImpl.validateLogEvent]: LogEvent.message='undefined' property must be a string."
    Assert.throws(
      () => LogEventImpl.validateLogEvent({ type: actualType, message: undefined as unknown as string, timestamp: actualtimeStamp }),
      ScriptError,
      errMsg,
      "LogEventImpl(validateLogEvent)-undefined message"
    )

    // Testing validateLogEvent (exception cases): null timestamp
    errMsg = "[LogEventImpl.validateLogEvent]: LogEvent.timestamp='null' property must be a Date."
    Assert.throws(
      () => LogEventImpl.validateLogEvent({ type: actualType, message: actualMsg, timestamp: null as unknown as Date }),
      ScriptError,
      errMsg,
      "LogEventImpl(validateLogEvent)-null timestamp"
    )

    // Testing validateLogEvent (exception cases): undefined timestamp
    errMsg = "[LogEventImpl.validateLogEvent]: LogEvent.timestamp='undefined' property must be a Date."
    Assert.throws(
      () => LogEventImpl.validateLogEvent({ type: actualType, message: actualMsg, timestamp: undefined as unknown as Date }),
      ScriptError,
      errMsg,
      "LogEventImpl(validateLogEvent)-undefined timestamp"
    )

    // Testing validateLogEvent (valid case)
    Assert.doesNotThrow(
      () => LogEventImpl.validateLogEvent({ type: actualType, message: actualMsg, timestamp: actualtimeStamp }),
      "LogEventImpl(validateLogEvent)-valid case"
    )

    // Testing validateLogEvent with extra fields (valid case)
    Assert.doesNotThrow(
      () => LogEventImpl.validateLogEvent({ type: actualType, message: actualMsg, timestamp: actualtimeStamp, extraFields: eventExtras }),
      "LogEventImpl(validateLogEvent)-valid case with extra fields(valid case"
    )
    // Testing validateLogEvent with extra fields (invalid case)
    errMsg = "[LogEventImpl.validateLogEvent]: extraFields must be a non-null plain object."
    Assert.throws(
      () => LogEventImpl.validateLogEvent({ type: actualType, message: actualMsg, timestamp: actualtimeStamp, extraFields: "invalid" as unknown as LogEventExtraFields }),
      ScriptError,
      errMsg,
      "LogEventImpl(validateLogEvent)-extraFields is a string-not a plain object"
    )
    // Testing validateLogEvent with extra fields (null case)
    errMsg = "[LogEventImpl.validateLogEvent]: extraFields must be a non-null plain object."
    Assert.throws(
      () => LogEventImpl.validateLogEvent({ type: actualType, message: actualMsg, timestamp: actualtimeStamp, extraFields: null as unknown as LogEventExtraFields }),
      ScriptError,
      errMsg,
      "LogEventImpl(validateLogEvent)-invalid case with extra fields-null"
    )

    // Testing validateLogEvent with extra fields (undefined valid case)
    // undefined is valid, since it defaults to an empty object
    Assert.doesNotThrow(
      () => LogEventImpl.validateLogEvent({ type: actualType, message: actualMsg, timestamp: actualtimeStamp, extraFields: undefined }),
      "LogEventImpl(validateLogEvent)-valid case with extra fields-undefined"
    )

    // ExtraFields: Reserved keys should not be included
    eventExtras = { type: "SHOULD NOT INCLUDE", message: "NO", timestamp: 123, toString: () => "no" }
    actualEvent = new LogEventImpl("msg", LOG_EVENT.INFO, eventExtras, new Date())
    Assert.equals(Object.keys(actualEvent.extraFields).length, 0, "LogEventImpl: Reserves keys: all are reserved in extraFields")
    eventExtras = { name: "Valid", message: "NO", timestamp: 123, toString: () => "no" }
    actualEvent = new LogEventImpl("msg", LOG_EVENT.INFO, eventExtras, new Date())
    Assert.equals(Object.keys(actualEvent.extraFields).length, 1, "LogEventImpl: Reserves keys: not all are reserved in extraFields")

    TestCase.clear()
  }

  public static consoleAppender(): void { // Unit Tests for ConsoleAppender class
    TestCase.clear()
    AbstractAppender.clearLogEventFactory() // In case other test case initialized it

    // No need to set the layout, since we are testing also default configuration of the AbstractAppender class

    // Defining the variables to be used in the tests
    let expectedStr: string, actualStr: string, expectedEvent: LogEvent,
      actualEvent: LogEvent | null, appender: Appender, layout: Layout, expectedNull: LogEvent | null,
      actualMsg: string, expectedMsg: string, msg: string, expectedType: LOG_EVENT, actualType: LOG_EVENT, errMsg: string,
      extraFields: LogEventExtraFields

    // Test lazy initialization: We can't because we need and instance first

    // Initial situation (testing information in AbstractAppender common to all appenders, no need to test it in each appender)
    appender = ConsoleAppender.getInstance()
    Assert.isNotNull(appender, "ConsoleAppender(getInstance) is not null")

    // Multiple calls to getInstance() returns the same object
    ConsoleAppender.clearInstance()
    const a1 = ConsoleAppender.getInstance()
    const a2 = ConsoleAppender.getInstance()
    Assert.equals(a1, a2, "ConsoleAppender(getInstance) - singleton reference test")

    // Testing static properties have default values (null)
    Assert.isNull((appender as AbstractAppender).getLastLogEvent(), "ConsoleAppender(getInstance) has no last log event")

    // To test the initial configuration we would need toString, since the getters do lazy initialization
    expectedStr = `AbstractAppender: {layout=null, logEventFactory="null", lastLogEvent=null} ConsoleAppender: {}`
    actualStr = (appender as ConsoleAppender).toString()
    Assert.equals(actualStr, expectedStr, "ConsoleAppender(getInstance) testing not initialized static properties from AbstractAppender via toString()")

    // Checking static properties
    layout = AbstractAppender.getLayout() // lazy initialized
    Assert.isNotNull(layout, "ConsoleAppender(getInstance) has a layout")

    // Checking the default formatter via LayoutImpl.toString()
    expectedStr = `LayoutImpl: {formatter: [Function: "defaultLayoutFormatterFun"]}`
    actualStr = (layout as LayoutImpl).toString()
    Assert.equals(actualStr, expectedStr, "ConsoleAppender(getInstance) has a default layout formatter")

    // Testing setting layout twice, second call should not change the layout
    AbstractAppender.clearLayout()
    const layout1 = new LayoutImpl()
    const layout2 = new LayoutImpl(LayoutImpl.shortFormatterFun)
    AbstractAppender.setLayout(layout1)
    AbstractAppender.setLayout(layout2) // Should NOT replace layout1
    Assert.equals(AbstractAppender.getLayout(), layout1, "AbstractAppender(setLayout) - second setLayout call does not override")
    AbstractAppender.clearLayout() // Reset the appender state
    AbstractAppender.setLayout(new LayoutImpl(LayoutImpl.defaultFormatterFun)) // Now set the second layout

    // setLayout: Invalid layout
    AbstractAppender.clearLayout()
    errMsg = `[AbstractAppender.setLayout]: Invalid Layout: layout object is null or undefined`
    Assert.throws(
      () => AbstractAppender.setLayout("not a layout" as unknown as Layout),
      ScriptError,
      errMsg,
      "AbstractAppender(setLayout) - invalid layout"
    )
    AbstractAppender.setLayout(new LayoutImpl()) // Reset the appender state to a valid layout

    // Setting log event factory twice (should not override the first)
    AbstractAppender.clearLogEventFactory()
    const factory1 = function f1(message:string, type:LOG_EVENT) { return new LogEventImpl("A" + message, type) }
    const factory2 = function f2(message:string, type:LOG_EVENT) { return new LogEventImpl("B" + message, type) }
    AbstractAppender.setLogEventFactory(factory1)
    AbstractAppender.setLogEventFactory(factory2) // Should NOT replace factory1
    Assert.equals(AbstractAppender.getLogEventFactory(), factory1, "AbstractAppender(setLogEventFactory) - second call does not override")
    AbstractAppender.clearLogEventFactory() // Reset the appender state
    AbstractAppender.setLogEventFactory(AbstractAppender.defaultLogEventFactoryFun) // Now set the second factory

    // Pass a non-function to setLogEventFactory and expect a ScriptError.
    AbstractAppender.clearLogEventFactory()
    errMsg = `[AbstractAppender.setLogEventFactory]: Invalid logEventFactory: Not a function`
    Assert.throws(
      () => AbstractAppender.setLogEventFactory("not a function" as unknown as LogEventFactory),
      ScriptError,
      errMsg,
      "AbstractAppender(setLogEventFactory) - invalid factory"
    )
    AbstractAppender.setLogEventFactory(AbstractAppender.defaultLogEventFactoryFun) // Now set the second factory

    // Testing log(LogEvent): valid case
    expectedMsg = "Info Message with lazy initialization testing ConsoleAppender"
    expectedType = LOG_EVENT.INFO
    expectedEvent = new LogEventImpl(expectedMsg, expectedType)
    Assert.doesNotThrow(() => appender!.log(expectedEvent),
      "ConsoleAppender(log(LogEvent)) - valid case with lazy initialization"
    )
    actualEvent = appender!.getLastLogEvent()
    Assert.equals(actualEvent?.type, expectedEvent?.type,
      "ConsoleAppender(getLastLogEvent not empty)-log(LogEvent).type with lazy initialization")
    Assert.equals(actualEvent?.message, expectedEvent?.message,
      "ConsoleAppender(getLastLogEvent not empty)-log(LogEvent).message with lazy initialization")

    // Testing log(LogEvent): valid case with extra fields
    extraFields = { userId: 123, sessionId: "abc" }
    expectedMsg = "Error Message with lazy initialization and extra fields testing ConsoleAppender"
    expectedType = LOG_EVENT.ERROR
    expectedEvent = new LogEventImpl(expectedMsg, expectedType, extraFields)
    Assert.doesNotThrow(() => appender!.log(expectedEvent),
      "ConsoleAppender(log(LogEvent)) - valid case with lazy initialization and extra fields"
    )
    actualEvent = appender!.getLastLogEvent()
    Assert.equals(actualEvent?.type, expectedEvent?.type,
      "ConsoleAppender(getLastLogEvent not empty)-log(LogEvent).type with lazy initialization and extra fields")
    Assert.equals(actualEvent?.message, expectedEvent?.message,
      "ConsoleAppender(getLastLogEvent not empty)-log(LogEvent).message with lazy initialization and extra fields")
    Assert.equals(actualEvent?.extraFields.userId, extraFields.userId,
      "ConsoleAppender(getLastLogEvent not empty)-log(LogEvent).extraFields.userId with lazy initialization and extra fields"
    )
    Assert.equals(actualEvent?.extraFields.sessionId, extraFields.sessionId,
      "ConsoleAppender(getLastLogEvent not empty)-log(LogEvent).extraFields.sessionId with lazy initialization and extra fields"
    )

    // Testing log(string, LOG_EVENT)
    expectedMsg = "Trace Message with lazy initialization testing ConsoleAppender"
    expectedType = LOG_EVENT.TRACE
    expectedEvent = new LogEventImpl(expectedMsg, expectedType)
    Assert.doesNotThrow(() => appender!.log(expectedMsg, expectedType),
      "ConsoleAppender(log(string,LOG_EVENT)) - valid case with lazy initialization"
    )
    actualEvent = appender!.getLastLogEvent()
    Assert.equals(actualEvent?.type, expectedEvent?.type,
      "ConsoleAppender(getLastMsg not empty)-log(string,LOG_EVENT).type")
    Assert.equals(actualEvent?.message, expectedEvent?.message,
      "ConsoleAppender(getLastLogEvent not empty)-log(string,LOG_EVENT).message with lazy initialization")

    // Testing log(string, LOG_EVENT) with extra fields
    extraFields = { userId: 456, sessionId: "xyz" }
    expectedMsg = "Trace Message with lazy initialization and extra fields testing ConsoleAppender"
    expectedType = LOG_EVENT.TRACE
    expectedEvent = new LogEventImpl(expectedMsg, expectedType, extraFields)
    Assert.doesNotThrow(() => appender!.log(expectedMsg, expectedType, extraFields),
      "ConsoleAppender(log(string,LOG_EVENT)) - valid case with lazy initialization and extra fields"
    )
    actualEvent = appender!.getLastLogEvent()
    Assert.equals(actualEvent?.type, expectedEvent?.type,
      "ConsoleAppender(getLastMsg not empty)-log(string,LOG_EVENT).type with lazy initialization and extra fields")
    Assert.equals(actualEvent?.message, expectedEvent?.message,
      "ConsoleAppender(getLastLogEvent not empty)-log(string,LOG_EVENT).message with lazy initialization and extra fields")
    Assert.equals(actualEvent?.extraFields.userId, extraFields.userId,
      "ConsoleAppender(getLastLogEvent not empty)-log(string,LOG_EVENT).extraFields.userId with lazy initialization and extra fields"
    )
    Assert.equals(actualEvent?.extraFields.sessionId, extraFields.sessionId,
      "ConsoleAppender(getLastLogEvent not empty)-log(string,LOG_EVENT).extraFields.sessionId with lazy initialization and extra fields"
    )

    // Testing log(LogEvent): null/undefined case
    errMsg = "[AbstractAppender.log]: LogEvent must be a non-null object."
    Assert.throws(
      () => appender!.log(null as unknown as LogEvent),
      ScriptError,
      errMsg,
      "ConsoleAppender(ScriptError)-log(LogEvent) - null"
    )
    Assert.throws(
      () => appender!.log(undefined as unknown as LogEvent),
      ScriptError,
      errMsg,
      "ConsoleAppender(ScriptError)-log(LogEvent) - undefined"
    )

    // Testing log(string, LOG_EVENT): null/undefined case

    // Testing after clearing the instance the last message is empty again
    ConsoleAppender.clearInstance()
    appender = ConsoleAppender.getInstance()
    actualEvent = appender.getLastLogEvent()
    Assert.equals(actualEvent, null,
      "ConsoleAppender(getLastMsg empty after clear the instance)")

    // Testing toString, default logEventFactory
    ConsoleAppender.clearInstance()
    expectedMsg = "Error message in appenders.toString testing ConsoleAppender"
    expectedType = LOG_EVENT.ERROR
    appender = ConsoleAppender.getInstance()
    actualEvent = new LogEventImpl(expectedMsg, expectedType)
    appender.log(actualEvent)
    actualStr = appender.toString()
    // Using the trick to call supper to string using prototype, to simplify the checking
    expectedStr = `${AbstractAppender.prototype.toString.call(appender)} ConsoleAppender: {}`
    Assert.equals(actualStr, expectedStr,
      "ConsoleAppender(toString)-default logEventFactory - toString matches expected format"
    )

    // Testing throwing a ScriptError: Not instanciated singleton and calling log method
    ConsoleAppender.clearInstance()
    appender = ConsoleAppender.getInstance()
    expectedStr = "[AbstractAppender.log]: event type='-1' must be provided and must be a valid LOG_EVENT value."
    Assert.throws(
      () => appender.log("Info event message", -1 as LOG_EVENT),
      ScriptError,
      expectedStr,
      "ConsoleAppender(ScriptError)-log - non valid LOG_EVENT"
    )

    // Testing throwing a ScriptError: Testing not instantiated singleton and calling toString method
    ConsoleAppender.clearInstance()
    expectedStr = "[ConsoleAppender.toString]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => appender.toString(),
      ScriptError,
      expectedStr,
      "ConsoleAppender(ScriptError)-toString-with singleton undefined"
    )

    // Testing NOT throwing a ScriptError: Testing not instantiated singleton and getting log event factory
    ConsoleAppender.clearInstance()
    Assert.doesNotThrow(
      () => ConsoleAppender.getLogEventFactory(),
      "ConsoleAppender(ScriptError)-getLogEventFactory - no custom factory"
    )

    // Testing custom Log event factory with a valid factory
    ConsoleAppender.clearInstance()
    const ENV = "PROD-"
    let prodLogEventFactoryFun: LogEventFactory
      = function logEventFactoryFun(message: string, eventType: LOG_EVENT) {
        return new LogEventImpl(ENV + message, eventType) // add environment prefix
      }
    appender = ConsoleAppender.getInstance()
    AbstractAppender.clearLogEventFactory()
    AbstractAppender.setLogEventFactory(prodLogEventFactoryFun)
    expectedMsg = "Custom LogEvent factory message testing ConsoleAppender"
    expectedType = LOG_EVENT.INFO
    expectedEvent = new LogEventImpl(ENV + expectedMsg, expectedType)
    appender.log(expectedMsg, expectedType)
    actualEvent = appender.getLastLogEvent()
    Assert.equals(actualEvent?.type, expectedEvent?.type,
      "ConsoleAppender(getLastMsg not empty)-log(string,LOG_EVENT) with custom factory.type")
    Assert.equals(actualEvent?.message, expectedEvent?.message,
      "ConsoleAppender(getLastMsg not empty)-log(string,LOG_EVENT) with custom factory.message")

    // Clear instance also clear the last log event
    appender.log("Some message", LOG_EVENT.INFO)
    ConsoleAppender.clearInstance()
    appender = ConsoleAppender.getInstance()
    Assert.equals(appender.getLastLogEvent(), null, "ConsoleAppender(clearInstance) - lastLogEvent is null after clear")

    AbstractAppender.clearLogEventFactory()

    TestCase.clear()
  }

  // Unit Tests for ExcelAppener claass
  public static excelAppender(workbook: ExcelScript.Workbook, msgCell: string): void {
    TestCase.clear()

    // Defining the variables to be used in the tests
    let expectedStr: string, actualStr: string, msg: string, expectedEvent: LogEvent,
      actualEvent: LogEvent | null, expectedMsg: string, expectedType: LOG_EVENT, errMsg: string,
      appender: Appender, msgCellRng: ExcelScript.Range, extraFields: LogEventExtraFields,
      activeSheet: ExcelScript.Worksheet, eventFonts: Record<LOG_EVENT, string>,
      layout: Layout

    activeSheet = workbook.getActiveWorksheet()
    msgCellRng = activeSheet.getRange(msgCell)
    const address = msgCellRng.getAddress()
    appender = ExcelAppender.getInstance(msgCellRng)
    const DEFAULT_FONTS = { ...ExcelAppender.DEFAULT_EVENT_FONTS }

    // Note: Testing log calls (may be redundant because log is from AbstractAppender, but we need to test the ExcelAppender specific behavior)
    // Testing sending log(message, LOG_EVENT)
    expectedMsg = "Info event testing ExcelAppender"
    expectedType = LOG_EVENT.INFO
    Assert.doesNotThrow(() => appender.log(expectedMsg, expectedType), "ExcelAppender(log(string,LOG_EVENT)) - valid case")
    actualStr = msgCellRng.getValue().toString()
    actualEvent = appender.getLastLogEvent()
    TestCase.runTestAsync(() => {
      Assert.isNotNull(actualEvent, "ExcelAppender(getLastLogEvent) not null")
      Assert.equals(actualEvent!.type, expectedType, "ExcelAppender(getLastLogEvent).type")
      Assert.equals(actualEvent!.message, expectedMsg, "ExcelAppender(getLastLogEvent).message")
      // Now checking the excel cell value (formatted via format method)
      expectedStr = AbstractAppender.getLayout().format(actualEvent as LogEvent)
      Assert.equals(actualStr, expectedStr, "ExcelAppender(cell value via log(string,LOG_EVENT))")
    })

    // Testing sending log(message, LOG_EVENT) with extra fields
    extraFields = { userId: 123, sessionId: "abc" }
    expectedMsg = "Info event with extra fields testing ExcelAppender"
    expectedType = LOG_EVENT.INFO
    Assert.doesNotThrow(() => appender.log(expectedMsg, expectedType, extraFields),
      "ExcelAppender(log(string,LOG_EVENT)) - valid case with extra fields"
    )
    TestCase.runTestAsync(() => {
      actualEvent = appender.getLastLogEvent()
      Assert.isNotNull(actualEvent, "ExcelAppender(getLastLogEvent) not null")
      Assert.equals(actualEvent!.type, expectedType, "ExcelAppender(getLastLogEvent).type with extra fields")
      Assert.equals(actualEvent!.message, expectedMsg, "ExcelAppender(getLastLogEvent).message with extra fields")
      Assert.equals(actualEvent!.extraFields.userId, extraFields.userId,
        "ExcelAppender(getLastLogEvent).extraFields.userId with extra fields"
      )
      Assert.equals(actualEvent!.extraFields.sessionId, extraFields.sessionId,
        "ExcelAppender(getLastLogEvent).extraFields.sessionId with extra fields"
      )
      // Now checking the excel cell value (formatted via format method)
      actualStr = msgCellRng.getValue().toString()
      expectedStr = AbstractAppender.getLayout().format(actualEvent as LogEvent)
      Assert.equals(actualStr, expectedStr, "ExcelAppender(cell value via log(message,LOG_EVENT)-with extra fields)")
    })

    // Testing log(LogEvent)
    actualEvent = appender.getLastLogEvent()
    expectedEvent = new LogEventImpl(expectedMsg, expectedType, {}, actualEvent.timestamp)
    Assert.doesNotThrow(() => appender.log(expectedEvent),
      "ExcelAppender(log(LogEvent)) - valid case"
    )
    actualEvent = appender.getLastLogEvent()
    actualStr = msgCellRng.getValue().toString()
    expectedStr = AbstractAppender.getLayout().format(expectedEvent as LogEvent)
    actualStr = msgCellRng.getValue().toString()
    Assert.equals(actualStr, expectedStr, "ExcelAppender(cell value) via log(LogEvent)")

    // Testing the corresponding last log event (redundant check)
    actualEvent = appender.getLastLogEvent()
    Assert.isNotNull(actualEvent, "ExcelAppender(getLastEvent) not null from log(LogEvent")
    Assert.equals(actualEvent!.type, expectedEvent.type,
      "ExcelAppender(getLastEvent).type from log(LogEvent")
    Assert.equals(actualEvent!.message, expectedEvent.message,
      "ExcelAppender(getLastEvent).message from log(LogEvent)")

    // Testing log(LogEvent) with extra fields
    expectedMsg = "Error event with extra fields testing ExcelAppender"
    expectedType = LOG_EVENT.ERROR
    expectedEvent = new LogEventImpl(expectedMsg, expectedType, extraFields)
    Assert.doesNotThrow(() => appender.log(expectedEvent),
      "ExcelAppender(log(LogEvent)) - valid case with extra fields"
    )
    TestCase.runTestAsync(() => {
      actualEvent = appender.getLastLogEvent()
      Assert.isNotNull(actualEvent, "ExcelAppender(getLastLogEvent) not null from log(LogEvent) with extra fields")
      Assert.equals(actualEvent!.type, expectedEvent.type,
        "ExcelAppender(getLastLogEvent).type from log(LogEvent) with extra fields")
      Assert.equals(actualEvent!.message, expectedEvent.message,
        "ExcelAppender(getLastLogEvent).message from log(LogEvent) with extra fields")
      Assert.equals(actualEvent!.extraFields.userId, extraFields.userId,
        "ExcelAppender(getLastLogEvent).extraFields.userId from log(LogEvent) with extra fields"
      )
      // Now checking the excel cell value (formatted via format method)
      actualStr = msgCellRng.getValue().toString()
      expectedStr = AbstractAppender.getLayout().format(actualEvent as LogEvent)
      Assert.equals(actualStr, expectedStr, "ExcelAppender(cell value via log(LogEvent)-with extra fields)")
    })

    // Custom log event valid colors
    ExcelAppender.clearInstance()
    Assert.doesNotThrow(
      () => ExcelAppender.getInstance(msgCellRng, null as unknown as Record<LOG_EVENT, string>),
      "ExcelAppender(getInstance)-using null log event fonts"
    )
    Assert.equals(
      (appender as ExcelAppender).getEventFonts(),
      DEFAULT_FONTS,
      "ExcelAppender(getEventFonts)-using null log event fonts defaults to DEFAULT_EVENT_FONTS"
    )

    // Changing two colors and the rest will be the defalt ones
    eventFonts = { ...DEFAULT_FONTS, [LOG_EVENT.WARN]: "#ed7d31", [LOG_EVENT.INFO]: "#548235" }
    ExcelAppender.clearInstance()
    appender = ExcelAppender.getInstance(msgCellRng, eventFonts)
    Assert.equals(
      (appender as ExcelAppender).getEventFonts(),
      eventFonts,
      "ExcelAppender(getEventFonts)-using custom log event fonts"
    )

    // Script Errors
    ExcelAppender.clearInstance() // singleton is undefined
    TestCase.runTestAsync(() => {
      errMsg = "[AbstractAppender.log]: A singleton instance can't be undefined or null. Please invoke getInstance first"
      Assert.throws(
        () => appender.log("Info message", LOG_EVENT.INFO),
        ScriptError,
        errMsg,
        "ExcelAppender(ScriptError)-log-singleton not defined"
      )
      // Script Errors: Testing non valid input: getInstancce(null)
      errMsg = "[ExcelAppender.getInstance]: A valid ExcelScript.Range for input argument msgCellRng is required."
      Assert.throws(
        () => ExcelAppender.getInstance(null),
        ScriptError,
        errMsg,
        "ExcelAppender(ScriptError)-getInstance(Non valid msgCellRng-null)"
      )
    })

    // Script Errors: Testing non valid input: getInstancce(undefined)
    Assert.throws(
      () => ExcelAppender.getInstance(undefined),
      ScriptError,
      errMsg,
      "ExcelAppender(ScriptError))-getInstance - Non valid msgCellRng-undefined"
    )

    // Script Errors: Testing non valid input: log with non valid LOG_EVENT
    appender = ExcelAppender.getInstance(activeSheet.getRange(msgCell))
    errMsg = "[AbstractAppender.log]: event type='-1' must be provided and must be a valid LOG_EVENT value."
    Assert.throws(
      () => appender.log("Info event message", -1 as LOG_EVENT),
      ScriptError,
      errMsg,
      "ExcelAppender(ScriptError)-Log non valid LOG_EVENT"
    )

    ExcelAppender.clearInstance()
    /*Mock object for ExcelScript.Range to simulate a multi-cell range in VS/TypeScript tests.
    This enables testing single-cell validation logic in environments where the real API isn't available.
    (Office Scripts allows any range; in VS strict typing and missing API require this manual mock.)*/
    const mockArrRng = { getCellCount: () => 2, setValue: () => { } }
    errMsg = "[ExcelAppender.getInstance]: Input argument msgCellRng must represent a single Excel cell."
    Assert.throws(
      () => ExcelAppender.getInstance(mockArrRng as unknown as ExcelScript.Range),
      ScriptError,
      errMsg,
      "ExcelAppender(ScriptError)-getInstance not a single cell"
    )

    // Script Errors: Testing non-valid hexadecimal colors
    ExcelAppender.clearInstance()
    errMsg = "[ExcelAppender.getInstance]: The input value '' for 'WARN' event is missing or not a string. Please provide a 6-digit hexadecimal color as 'RRGGBB' or '#RRGGBB'."
    eventFonts = { ...DEFAULT_FONTS, [LOG_EVENT.WARN]: "" }
    Assert.throws(
      () => ExcelAppender.getInstance(msgCellRng, eventFonts),
      ScriptError,
      errMsg,
      "ExcelAppender(ScriptError)-getInstance-Non valid font color for warning"
    )
    errMsg = "[ExcelAppender.getInstance]: The input value 'xxxxxx' for 'INFO' event is not a valid 6-digit hexadecimal color. Please use 'RRGGBB' or '#RRGGBB' format."
    eventFonts = { ...DEFAULT_FONTS, [LOG_EVENT.INFO]: "xxxxxx" }
    Assert.throws(
      () => ExcelAppender.getInstance(msgCellRng, eventFonts),
      ScriptError,
      errMsg,
      "ExcelAppender(ScriptError) - getInstance - Non valid font color for info"
    )
    errMsg = "[ExcelAppender.getInstance]: The input value '******' for 'TRACE' event is not a valid 6-digit hexadecimal color. Please use 'RRGGBB' or '#RRGGBB' format."
    eventFonts = { ...DEFAULT_FONTS, [LOG_EVENT.TRACE]: "******" }
    Assert.throws(
      () => ExcelAppender.getInstance(msgCellRng, eventFonts),
      ScriptError,
      errMsg,
      "ExcelAppender(ScriptError)-getInstance-Non valid font color for trace"
    )

    // Testing toString
    ExcelAppender.clearInstance()
    appender = ExcelAppender.getInstance(activeSheet.getRange(msgCell))
    msg = "Trace message in ExcelAppender toString"
    appender.log(msg, LOG_EVENT.TRACE)
    expectedStr = `${AbstractAppender.prototype.toString.call(appender)}`
    // address in the expected string is dynamic and works cross` platform (Office Script and TypeScript)
    expectedStr += ` ExcelAppender: {msgCellRng(address)="${address}", eventfonts={ERROR="9c0006",WARN="ed7d31",INFO="548235",TRACE="7f7f7f"}}`
    actualStr = (appender as ExcelAppender).toString()
    Assert.equals(actualStr, expectedStr, "ExcelAppender(toString)")

    TestCase.clear()
  }

  public static loggerImpl(workbook: ExcelScript.Workbook, msgCell: string): void { // Unit tests for LoggerImpl class
    TestCase.clear()
    // Defining variables
    let logger: Logger, actualStr: string, expectedStr: string, appender: Appender

    // Checking Initial situation
    logger = LoggerImpl.getInstance()
    Assert.isNotNull(logger, "LoggerImpl(getInstance) is not null")
    Assert.equals(logger!.getLevel(), LoggerImpl.LEVEL.WARN, "LoggerImpl(getInstance)-default level is WARN")
    Assert.equals(logger!.getAction(), LoggerImpl.ACTION.EXIT, "LoggerImpl(getInstance)-default action is EXIT")
    Assert.isNotNull(logger!.getAppenders(), "LoggerImpl(getInstance)-default appenders is not null")
    Assert.equals(logger!.getAppenders().length, 0, "LoggerImpl(getInstance)-default appenders length is 0")

    // Testing getting label for LEVEL and ACTION
    expectedStr = "OFF"
    actualStr = LoggerImpl.getLevelLabel(LoggerImpl.LEVEL.OFF)
    Assert.equals(actualStr, expectedStr, "LoggerImpl(getLevelLabel)-OFF label is correct")
    expectedStr = "WARN" // Default level
    actualStr = LoggerImpl.getLevelLabel(undefined) // non valid level
    Assert.equals(actualStr, expectedStr, "LoggerImpl(getLevelLabel)-non valid level label is WARN")

    // Testing getting action label
    expectedStr = "CONTINUE"
    actualStr = LoggerImpl.getActionLabel(LoggerImpl.ACTION.CONTINUE)
    Assert.equals(actualStr, expectedStr, "LoggerImpl(getActionLabel)-CONTINUE label is correct")
    expectedStr = "EXIT" // Default action
    actualStr = LoggerImpl.getActionLabel(undefined) // non valid action
    Assert.equals(actualStr, expectedStr, "LoggerImpl(getActionLabel)-non valid action label is UNKNOWN")

    // Testing adding/removing appenders
    appender = ConsoleAppender.getInstance()
    Assert.doesNotThrow(() => logger.addAppender(appender), "LoggerImpl(addAppender) - valid case")
    Assert.equals(logger.getAppenders().length, 1, "LoggerImpl(addAppender) - appender added")
    Assert.isTrue(logger.getAppenders().includes(appender), "LoggerImpl(addAppender) - appender is in the list")
    Assert.doesNotThrow(() => logger.removeAppender(appender), "LoggerImpl(removeAppender) - valid case")
    Assert.equals(logger.getAppenders().length, 0, "LoggerImpl(removeAppender) - appender removed")
    Assert.isFalse(logger.getAppenders().includes(appender), "LoggerImpl(removeAppender) - appender is not in the list")
    Assert.doesNotThrow(() => logger.removeAppender(appender), "LoggerImpl(removeAppender) - empty list valid case")
    Assert.equals(logger.getAppenders().length, 0, "LoggerImpl(removeAppender) - empty list valid case")

    TestCase.clear()

    // Testing scenario based on different combinations of LEVEL and ACTION
    // It creates all log event for each combination of LEVEL and ACTION
    // ACTION.CONTINUE
    TestCase.loggerImplLevels(false, LoggerImpl.LEVEL.OFF, LoggerImpl.ACTION.CONTINUE, workbook, msgCell)
    TestCase.loggerImplLevels(false, LoggerImpl.LEVEL.ERROR, LoggerImpl.ACTION.CONTINUE, workbook, msgCell)
    TestCase.loggerImplLevels(false, LoggerImpl.LEVEL.WARN, LoggerImpl.ACTION.CONTINUE, workbook, msgCell)
    TestCase.loggerImplLevels(false, LoggerImpl.LEVEL.INFO, LoggerImpl.ACTION.CONTINUE, workbook, msgCell)
    TestCase.loggerImplLevels(false, LoggerImpl.LEVEL.TRACE, LoggerImpl.ACTION.CONTINUE, workbook, msgCell)
    // ACTION.EXIT
    TestCase.loggerImplLevels(false, LoggerImpl.LEVEL.OFF, LoggerImpl.ACTION.EXIT, workbook, msgCell)
    TestCase.loggerImplLevels(false, LoggerImpl.LEVEL.ERROR, LoggerImpl.ACTION.EXIT, workbook, msgCell)
    TestCase.loggerImplLevels(false, LoggerImpl.LEVEL.WARN, LoggerImpl.ACTION.EXIT, workbook, msgCell)
    TestCase.loggerImplLevels(false, LoggerImpl.LEVEL.INFO, LoggerImpl.ACTION.EXIT, workbook, msgCell)
    TestCase.loggerImplLevels(false, LoggerImpl.LEVEL.TRACE, LoggerImpl.ACTION.EXIT, workbook, msgCell)

    // Now considering extra fields
    // ACTION.CONTINUE
    TestCase.loggerImplLevels(true, LoggerImpl.LEVEL.OFF, LoggerImpl.ACTION.CONTINUE, workbook, msgCell)
    TestCase.loggerImplLevels(true, LoggerImpl.LEVEL.ERROR, LoggerImpl.ACTION.CONTINUE, workbook, msgCell)
    TestCase.loggerImplLevels(true, LoggerImpl.LEVEL.WARN, LoggerImpl.ACTION.CONTINUE, workbook, msgCell)
    TestCase.loggerImplLevels(true, LoggerImpl.LEVEL.INFO, LoggerImpl.ACTION.CONTINUE, workbook, msgCell)
    TestCase.loggerImplLevels(true, LoggerImpl.LEVEL.TRACE, LoggerImpl.ACTION.CONTINUE, workbook, msgCell)
    // ACTION.EXIT
    TestCase.loggerImplLevels(true, LoggerImpl.LEVEL.OFF, LoggerImpl.ACTION.EXIT, workbook, msgCell)
    TestCase.loggerImplLevels(true, LoggerImpl.LEVEL.ERROR, LoggerImpl.ACTION.EXIT, workbook, msgCell)
    TestCase.loggerImplLevels(true, LoggerImpl.LEVEL.WARN, LoggerImpl.ACTION.EXIT, workbook, msgCell)
    TestCase.loggerImplLevels(true, LoggerImpl.LEVEL.INFO, LoggerImpl.ACTION.EXIT, workbook, msgCell)
    TestCase.loggerImplLevels(true, LoggerImpl.LEVEL.TRACE, LoggerImpl.ACTION.EXIT, workbook, msgCell)

    TestCase.clear()
  }

  public static loggerImplLazyInit() { // Unit Tests on Lazy Initialization for Logger class (instance and appender)
    TestCase.clear()

    // Defining the variables to be used in the tests
    let expectedMsg: string, expectedType: LOG_EVENT, expectedEvent: LogEvent, logger: Logger,
      expectedNum: number, actualNum: number, actualEvent: LogEvent | null

    // Testing lazy initialization of the appender
    expectedMsg = "Info event, testing lazyInit for LoggerImpl"
    expectedType = LOG_EVENT.INFO
    // No appender was defined
    logger = LoggerImpl.getInstance(LoggerImpl.LEVEL.INFO) // initialized the singleton
    logger.info(expectedMsg) // lazy initialization of the appender
    expectedNum = 1
    actualNum = logger.getAppenders().length ?? 0
    Assert.equals(actualNum, expectedNum, "Logger(Lazy init)-appender size is one")
    Assert.isNotNull(logger.getAppenders()[0], "Logger(Lazy init)-appender is not null")
    Assert.equals(logger.getLevel(), LoggerImpl.LEVEL.INFO, "Logger(Lazy init)-level is INFO")
    Assert.equals(logger.getAction(), LoggerImpl.ACTION.EXIT, "Logger(Lazy init)-action is EXIT(default")
    actualEvent = logger.getAppenders()[0].getLastLogEvent() // Safe to use getLastLogEvent here since it was tested in ConsoleAppender
    Assert.isNotNull(actualEvent, "Logger(Lazy init)-getLastLogEvent is not null")
    Assert.equals(actualEvent.type, expectedType, "Logger(Lazy init)-getLastLogEvent.type is INFO")
    Assert.equals(actualEvent.message, expectedMsg, "Logger(Lazy init)-getLastLogEvent.message info message is correct")

    // Lazy initialization of the singleton with default parameters (WARN,EXIT)
    expectedMsg = "Error event testing lazyInit for LoggerImpl"
    expectedType = LOG_EVENT.ERROR
    expectedEvent = new LogEventImpl(expectedMsg, LOG_EVENT.ERROR)
    LoggerImpl.clearInstance()
    TestCase.runTestAsync(() => {
      Assert.isNotNull(LoggerImpl.getInstance(), "Lazy init(logger != null)")
      logger = LoggerImpl.getInstance() // Get alreazy lazy initialized singleton
      Assert.equals(logger.getLevel(), LoggerImpl.LEVEL.WARN, "Logger(Lazy init)-level is WARN")
      Assert.equals(logger.getAction(), LoggerImpl.ACTION.EXIT, "Logger(Lazy init)-action is EXIT")
    })

    // To check the ScriptError message, since it may include the timestamp, we would need to use a short layout
    TestCase.setShortLayout()
    Assert.throws(
      () => logger.error(expectedMsg),
      ScriptError,
      AbstractAppender.getLayout().format(expectedEvent as LogEvent), // regardless of the layout it works
      "Logger(Lazy init)-error event with appender lazy initialized and expected to throw ScriptError"
    )
    actualEvent = logger.getAppenders()[0].getLastLogEvent() // Safe to use getLastLogEvent here since it was tested in ConsoleAppender
    Assert.isNotNull(actualEvent, "Logger(Lazy init)-getLastLogEvent is not null")
    Assert.equals(actualEvent.type, expectedType, "Logger(Lazy init)-getLastLogEvent.type is ERROR")
    Assert.equals(actualEvent.message, expectedMsg, "Logger(Lazy init)-getLastLogEvent.message error message is correct")
    TestCase.setDefaultLayout()

    // Testing ScriptError when no singleton is defined
    LoggerImpl.clearInstance()

    // Singleton will be lazy initialized in log private method
    TestCase.setShortLayout()
    Assert.throws(
      () => logger.error(expectedMsg),
      ScriptError,
      AbstractAppender.getLayout().format(expectedEvent as LogEvent), // regardless of the layout it works
      "Logger(Lazy init)-singleton not initialized"
    )
    TestCase.setDefaultLayout()

    TestCase.clear()
  }

  /**Unit tests for Logger class checking the behaviour after the singleton was reset
   */
  public static loggerImplResetSingleton(workbook: ExcelScript.Workbook, msgCell: string): void {
    TestCase.clear()

    // Defining the variables to be used in the tests
    let logger: Logger, expectedErrMsg: string

    logger = LoggerImpl.getInstance()
    LoggerImpl.clearInstance() // Singleton is undefined
    // Now we need to invoke a method that doesn't invoke lazy initialization
    expectedErrMsg = "[LoggerImpl.getErrCnt]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.getErrCnt(),
      ScriptError,
      expectedErrMsg,
      "loggerResetSingleton(getErrCnt())"
    )
    expectedErrMsg = "[LoggerImpl.getWarnCnt]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.getWarnCnt(),
      ScriptError,
      expectedErrMsg,
      "loggerResetSingleton(getWarnCnt())"
    )
    expectedErrMsg = "[LoggerImpl.hasErrors]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.hasErrors(),
      ScriptError,
      expectedErrMsg,
      "loggerResetSingleton(hasErrors())"
    )
    expectedErrMsg = "[LoggerImpl.hasWarnings]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.hasWarnings(),
      ScriptError,
      expectedErrMsg,
      "loggerResetSingleton(hasWarnings())"
    )
    expectedErrMsg = "[LoggerImpl.getCriticalEvents]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.getCriticalEvents(),
      ScriptError,
      expectedErrMsg,
      "loggerResetSingleton(getMessages())"
    )
    expectedErrMsg = "[LoggerImpl.getAppenders]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.getAppenders(),
      ScriptError,
      expectedErrMsg,
      "loggerResetSingleton(getAppenders())"
    )
    expectedErrMsg = "[LoggerImpl.getAction]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.getAction(),
      ScriptError,
      expectedErrMsg,
      "loggerResetSingleton(getAction())"
    )
    expectedErrMsg = "[LoggerImpl.getLevel]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.getLevel(),
      ScriptError,
      expectedErrMsg,
      "loggerResetSingleton(getLevel())"
    )

    TestCase.clear()
  }

  /**Unit Tests for Logger class for testing counters */
  public static loggerImplCounters(): void {
    TestCase.clear()

    // Defining the variables to be used in the tests
    let logger: Logger, layout: Layout, actualNum: number, expectedNum: number,
      actualEvent: LogEvent | null, expectedEvent: LogEvent, errMsg: string, warnMsg: string

    // Initializing the logger with a short layout
    layout = new LayoutImpl(LayoutImpl.shortFormatterFun) // Short layout

    logger = LoggerImpl.getInstance(LoggerImpl.LEVEL.INFO, LoggerImpl.ACTION.CONTINUE)
    logger.addAppender(ConsoleAppender.getInstance())
    // Testing counters on initial state
    Assert.equals(logger.getErrCnt(), 0, "loggerCounters(getErrCnt=0)")
    Assert.equals(logger.getWarnCnt(), 0, "loggerCounters(getWarnCnt=0)")
    Assert.equals(logger.hasErrors(), false, "loggerCounters(hasErrors=false)")
    Assert.equals(logger.hasWarnings(), false, "loggerCounters(hasWarnings=false)")
    Assert.equals(logger.getCriticalEvents(), [], "loggerCounters(getMessages=[])")

    // Sending events affecting the counter
    errMsg = "Error event for testing counters for LoggerImpl"
    logger.error(errMsg)
    expectedNum = 1
    actualNum = logger.getErrCnt()
    Assert.equals(actualNum, expectedNum, "loggerCounters(getErrCnt=1)")
    actualEvent = logger.getCriticalEvents()[0] ?? null // Get the first event
    expectedEvent = new LogEventImpl(errMsg, LOG_EVENT.ERROR)
    Assert.isNotNull(actualEvent, "LoggerCounters(getMessage()[0]) not null")
    Assert.equals((actualEvent as LogEvent).type, expectedEvent.type, "LoggerCounters(getMessage()[0]).type")
    Assert.equals((actualEvent as LogEvent).message, expectedEvent.message, "LoggerCounters(getMessage()[0]).message")


    // Testing counter for warnings
    warnMsg = "Warning event for testing counters for LoggerImpl"
    logger.warn(warnMsg)
    expectedNum = 1
    actualNum = logger.getWarnCnt()
    Assert.equals(actualNum, expectedNum, "loggerCounters(getWarnCnt=1)")
    // Testing messages
    let expectedArr = TestCase.simplifyLogEvents([new LogEventImpl(errMsg, LOG_EVENT.ERROR),
    new LogEventImpl(warnMsg, LOG_EVENT.WARN)])
    let actualArr = TestCase.simplifyLogEvents(logger.getCriticalEvents())
    Assert.equals(actualArr, expectedArr, "loggerCounters(getMessages)")
    Assert.equals(logger.hasMessages(), true, "loggerCounters(hasMessages)")
    // Testing other events, don't affect the counters
    let msg = "Info event doesn't count testing counters for LoggerImpl"
    logger.info(msg)
    actualNum = logger.getErrCnt()
    Assert.equals(actualNum, expectedNum, "LoggerCounter(getErrCnt=1)")
    actualNum = logger.getWarnCnt()
    Assert.equals(actualNum, expectedNum, "LoggerCounter(getWarnCnt=1)")
    Assert.equals(actualArr, expectedArr, "loggerCounters(getMessages-2nd time)")

    // Clearing counts
    logger.reset()
    expectedNum = 0
    actualNum = logger.getErrCnt()
    Assert.equals(actualNum, expectedNum, "LoggerCounter(errors cleared)")
    actualNum = logger.getWarnCnt()
    Assert.equals(actualNum, expectedNum, "LoggerCounter(warnings cleared)")
    expectedArr = []
    actualArr = logger.getCriticalEvents()
    Assert.equals(actualArr, expectedArr, "LoggerCounter(messages cleared)")
    // Checking appenders were not removed
    expectedNum = 1
    actualNum = logger.getAppenders().length
    Assert.equals(actualNum, expectedNum, "LoggerCounter(appenders not removed)")
    TestCase.clear()
  }

  /**Unit Tests for Logger class on toString method */
  public static loggerImplToString(workbook: ExcelScript.Workbook, msgCell: string): void {
    TestCase.clear()
    TestCase.setShortLayout()
    // Defining the variables to be used in the tests
    let expected: string, actual: string, layout: Layout, logger: Logger, extraFields: LogEventExtraFields

    //layout = new LayoutImpl(LayoutImpl.shortFormatterFun) // Short layout
    logger = LoggerImpl.getInstance(LoggerImpl.LEVEL.INFO, // Level of verbose
      LoggerImpl.ACTION.CONTINUE)
    // Adding appenders
    logger.addAppender(ConsoleAppender.getInstance())
    logger.addAppender(ExcelAppender.getInstance(workbook.getActiveWorksheet().getRange(msgCell)))

    // Testing toString method
    const MSGS = ["Error event testing for loggerImpl.toString",
      "Warning event for loggerImpl.toString"]
    logger.error(MSGS[0]) // lazy initialization of the appender
    logger.warn(MSGS[1])
    expected = `LoggerImpl: {level: "INFO", action: "CONTINUE", errCnt: 1, warnCnt: 1, appenders: ` +
      `[AbstractAppender: {layout=LayoutImpl: {formatter: [Function: "shortLayoutFormatterFun"]}, ` +
      `logEventFactory="defaultLogEventFactoryFun", lastLogEvent=LogEventImpl: {timestamp="2025-06-18 01:02:39,720", ` +
      `type="WARN", message="Warning event for loggerImpl.toString"}} ConsoleAppender: {}, AbstractAppender: {layout=LayoutImpl: ` +
      `{formatter: [Function: "shortLayoutFormatterFun"]}, logEventFactory="defaultLogEventFactoryFun", lastLogEvent=LogEventImpl: ` +
      `{timestamp="2025-06-18 01:02:39,720", type="WARN", message="Warning event for loggerImpl.toString"}} ExcelAppender: ` +
      `{msgCellRng(address)="C2", eventfonts={ERROR="9c0006",WARN="ed7d31",INFO="548235",TRACE="7f7f7f"}}]}`
    actual = logger.toString()
    TestCase.runTestAsync(() => {
      Assert.equals(normalizeTimestamps(actual), normalizeTimestamps(expected), "loggerToString(Logger)")
    })

    // Testing toString with extra fields
    extraFields = { userId: 123, sessionId: "abc" }
    TestCase.clear()
    logger = LoggerImpl.getInstance(LoggerImpl.LEVEL.INFO)
    logger.info("Info event testing logger.toString with extra fields", extraFields)
    //console.log(`logger=${logger.toString()}`)
    expected = `LoggerImpl: {level: "INFO", action: "EXIT", errCnt: 0, warnCnt: 0, appenders: [AbstractAppender: {layout=LayoutImpl: ` +
      `{formatter: [Function: "defaultLayoutFormatterFun"]}, logEventFactory="defaultLogEventFactoryFun", ` +
      `lastLogEvent=LogEventImpl: {timestamp="2025-06-19 22:31:17,324", type="INFO", message="Info event testing logger.toString with extra fields", ` +
      `extraFields={"userId":123,"sessionId":"abc"}}} ConsoleAppender: {}]}`
    actual = logger.toString()
    Assert.equals(normalizeTimestamps(actual), normalizeTimestamps(expected), "loggerToString(Logger with extra fields)")

    // Testing shortToString method
    expected = `LoggerImpl: {level: "INFO", action: "EXIT", errCnt: 0, warnCnt: 0, appenders: [ConsoleAppender]}`
    actual = (logger as LoggerImpl).toShortString()
    Assert.equals(actual, expected, "loggerToString(LoggerImpl) short version")

    TestCase.clear()

    // Helper function to normalize timestamps in the expected and actual strings
    function normalizeTimestamps(str: string): string {
      return str.replace(/timestamp="[^"]*"/g, 'timestamp="<TIMESTAMP>"')
    }
  }

  /**Unit Tests for Logger class for method exportState */
  public static loggerImplExportState(): void {
    TestCase.clear()
    TestCase.setShortLayout()

    // Defining the variables to be used in the tests
    let logger: Logger, layout: Layout, expectedEvents: LogEvent[],
      messages: string[], state: {
        level: string, action: string,
        errorCount: number, warningCount: number, criticalEvents: LogEvent[]
      }, msgs: string[]

    layout = new LayoutImpl(LayoutImpl.shortFormatterFun) // Short layout
    logger = LoggerImpl.getInstance(LoggerImpl.LEVEL.TRACE, LoggerImpl.ACTION.CONTINUE)
    msgs = ["warning event for testing exportState for LoggerImpl",
      "error event for testing exportState for LoggerImpl"]
    logger.trace("trace event for testing exportState for LoggerImpl")
    logger.info("info event for testing exportState for LoggerImpl")
    logger.warn(msgs[0])
    logger.error(msgs[1])
    state = logger.exportState()
    Assert.equals(state.level, "TRACE", "loggerExportState(level)")
    Assert.equals(state.action, "CONTINUE", "loggerExportState(action)")
    Assert.equals(state.errorCount, 1, "loggerExportState(errorCount)")
    Assert.equals(state.warningCount, 1, "loggerExportState(warningCount)")
    Assert.equals(state.criticalEvents.length, 2, "loggerExportState(messages.length)")
    expectedEvents = [
      new LogEventImpl(msgs[0], LOG_EVENT.WARN),
      new LogEventImpl(msgs[1], LOG_EVENT.ERROR)]
    Assert.equals(TestCase.simplifyLogEvents(state.criticalEvents),
      TestCase.simplifyLogEvents(expectedEvents), "loggerExportState(messages))")
    TestCase.clear()
  }

  /**Unit tests for Logger class, for testing scenarios where a ScriptError will be thrown.
   * It also tests all defensing programming scenarios implemented.
   */
  public static loggerImplScriptError(workbook: ExcelScript.Workbook, msgCell: string) {
    TestCase.clear()
    TestCase.setShortLayout()

    // Defining the variables to be used in the tests
    let logger: Logger, expectedMsg: string, actualMsg: string, appender: Appender,
      consoleAppender: ConsoleAppender, excelAppender: ExcelAppender, activeSheet: ExcelScript.Worksheet,
      msgCellRng: ExcelScript.Range

    // Testing non valid Logger.ACTION
    expectedMsg = "[LoggerImpl.getInstance]: The input value level='-1', was not defined in Logger.LEVEL."
    Assert.throws(
      () => LoggerImpl.getInstance(-1, LoggerImpl.ACTION.CONTINUE),
      ScriptError,
      expectedMsg,
      "loggerScriptError-Non valid LOG_LEVEL enum value"
    )

    // Testing when is invoked validateInstance method
    LoggerImpl.clearInstance()
    logger = LoggerImpl.getInstance()
    LoggerImpl.clearInstance() // now _instance is undefined
    expectedMsg = "[LoggerImpl.getErrCnt]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.getErrCnt(),
      ScriptError,
      expectedMsg,
      "loggerScriptError(getErrCnt())"
    )
    expectedMsg = "[LoggerImpl.getWarnCnt]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.getWarnCnt(),
      ScriptError,
      expectedMsg,
      "loggerScriptError-(getWarnCnt())"
    )
    expectedMsg = "[LoggerImpl.getCriticalEvents]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.getCriticalEvents(),
      ScriptError,
      expectedMsg,
      "loggerScriptError-(getMessages())"
    )
    expectedMsg = "[LoggerImpl.getLevel]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.getLevel(),
      ScriptError,
      expectedMsg,
      "loggerScriptError(getLevel())"
    )
    expectedMsg = "[LoggerImpl.getAction]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.getAction(),
      ScriptError,
      expectedMsg,
      "loggerScriptError(getAction())"
    )
    expectedMsg = "[LoggerImpl.hasErrors]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.hasErrors(),
      ScriptError,
      expectedMsg,
      "loggerScriptError(hasErrors())"
    )
    expectedMsg = "[LoggerImpl.hasWarnings]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.hasWarnings(),
      ScriptError,
      expectedMsg,
      "loggerScriptError(hasWarnings())"
    )
    expectedMsg = "[LoggerImpl.clear]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.reset(),
      ScriptError,
      expectedMsg,
      "loggerScriptError(clear())"
    )
    // Testing adding/setting/removing appender with undefined/null singleton
    expectedMsg = "[LoggerImpl.getAppenders]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    consoleAppender = ConsoleAppender.getInstance()
    Assert.throws(
      () => logger.getAppenders(),
      ScriptError,
      expectedMsg,
      "loggerScriptError(getAppenders())-undefined singleton"
    )
    expectedMsg = "[LoggerImpl.addAppender]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.addAppender(consoleAppender),
      ScriptError,
      expectedMsg,
      "loggerScriptError(addAppender())-undefined singleton"
    )
    expectedMsg = "[LoggerImpl.removeAppender]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.removeAppender(consoleAppender),
      ScriptError,
      expectedMsg,
      "loggerScriptError(removeAppender())-undefined singleton"
    )
    expectedMsg = "[LoggerImpl.setAppenders]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.setAppenders([consoleAppender, consoleAppender]),
      ScriptError,
      expectedMsg,
      "loggerScriptError(setAppenders(duplicated))-undefined singleton"
    )
    expectedMsg = "[LoggerImpl.toString]: A singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.toString(),
      ScriptError,
      expectedMsg,
      "loggerScriptError(toString())"
    )
    LoggerImpl.clearInstance()
    // Testing adding a null/undefined appender to a valid singleton

    consoleAppender = ConsoleAppender.getInstance()
    LoggerImpl.clearInstance()
    LoggerImpl.getInstance()

    expectedMsg = "[LoggerImpl.addAppender]: You can't add an appender that is null or undefined"
    Assert.throws(
      () => logger.addAppender(null as unknown as Appender),
      ScriptError,
      expectedMsg,
      "loggerScriptError(addAppenders()-null-valid-singleton)"
    )
    Assert.throws(
      () => logger.addAppender(undefined as unknown as Appender),
      ScriptError,
      expectedMsg,
      "loggerScriptError(addAppenders()-undefined-valid-singleton)"
    )
    // Adding appenders via setAppenders
    expectedMsg = "[LoggerImpl.setAppenders]: Invalid input: the input argument 'appenders' must be a non-null array."
    Assert.throws(
      () => logger.setAppenders(undefined as unknown as Appender[]),
      ScriptError,
      expectedMsg,
      `Internal Error(setAppenders)-undefined-valid singleton`
    )
    Assert.throws(
      () => logger.setAppenders(null as unknown as Appender[]),
      ScriptError,
      expectedMsg,
      "loggerScriptError(setAppenders)-null-valid-singleton"
    )

    expectedMsg = "[LoggerImpl.setAppenders]: Input argument appenders array contains null or undefined entry."
    Assert.throws(
      () => logger.setAppenders([consoleAppender, null as unknown as Appender]),
      ScriptError,
      expectedMsg,
      "loggerScriptError-[consoleAppender,null]-valid singleton"
    )
    Assert.throws(
      () => logger.setAppenders([consoleAppender, undefined as unknown as Appender]),
      ScriptError,
      expectedMsg,
      "loggerScriptError-[consoleAppender,undefined]-valid singleton"
    )
    expectedMsg = "[LoggerImpl.setAppenders]: Only one appender of type ConsoleAppender is allowed."
    Assert.throws(
      () => logger.setAppenders([consoleAppender, consoleAppender]),
      ScriptError,
      expectedMsg,
      "loggerScriptError-[consoleAppender,consoleAppender]-valid singleton"
    )

    // Testing adding duplicate appender
    LoggerImpl.clearInstance()
    logger = LoggerImpl.getInstance()
    logger.addAppender(ConsoleAppender.getInstance())
    expectedMsg = "[LoggerImpl.addAppender]: Only one appender of type ConsoleAppender is allowed."
    Assert.throws(
      () => logger.addAppender(ConsoleAppender.getInstance()),
      ScriptError,
      expectedMsg,
      "loggerScriptError-addaAppender duplicated"
    )
    LoggerImpl.clearInstance()
    LoggerImpl.getInstance()
    activeSheet = workbook.getActiveWorksheet()
    excelAppender = ExcelAppender.getInstance(activeSheet.getRange(msgCell))
    expectedMsg = "[LoggerImpl.setAppenders]: Only one appender of type ExcelAppender is allowed."
    Assert.throws(
      () => logger.setAppenders([excelAppender, excelAppender]),
      ScriptError,
      expectedMsg,
      "loggerScriptError-setAppender - duplicated"
    )

    TestCase.clear()
  }

}

// ----------------------------------------
// End Testing the Logging framework
// ----------------------------------------
// Make main available globally for Node/ts-node test environments
if (typeof globalThis !== "undefined" && typeof main !== "undefined") {
  // @ts-ignore
  globalThis.main = main
}
// #endregion main.ts

