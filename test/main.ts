// ----------------------------------------
// Testing the Logging framework
// ----------------------------------------

//Main function of the Script
function main(workbook: ExcelScript.Workbook,
) {

  // Parameters and constants definitions
  // ------------------------------------
  const MSG_CELL = "C2" // relative to the active sheet
  //const VERBOSITY = TestRunner.VERBOSITY.OFF        // unmment the scenario of your preference
  //const VERBOSITY = TestRunner.VERBOSITY.HEADER
  const VERBOSITY = TestRunner.VERBOSITY.SECTION
  const START_TEST = "START TEST"
  const END_TEST = "END TEST"
  const SHOW_TRACE = true

  let run: TestRunner = new TestRunner(VERBOSITY) // Controles the test execution process
  let success = false // Control variable to send the last message in finally

  // MAIN EXECUTION
  // --------------------

  try {
    const VERBOSITY_LEVEL = run.getVerbosityLabel()
    run.title(`${START_TEST} with verbosity '${VERBOSITY_LEVEL}'`, 1)
    let indent: number = 2 // Use the same indentation level for all test cases
    // Setting a commong layout

    /*All functions need to be invoked using arrow function (=>).
    Test cases organized by topics. They don't have any dependency, so they can
    be executed in any order.*/

    run.exec("Test Case ScriptError", () => TestCase.scriptError(), indent)
    run.exec("Test Case LayoutImpl", () => TestCase.testLayoutImpl(), indent)
    run.exec("Test Case LogEventImpl", () => TestCase.testLogEventImpl(), indent)
    run.exec("Test Case ConsoleAppender", () => TestCase.consoleAppender(), indent)
    run.exec("Test Case ExcelAppender", () => TestCase.excelAppender(workbook, MSG_CELL), indent)
    run.exec("Test Case Logger: Lazy Init", () => TestCase.loggerLazyInit(), indent)
    run.exec("Test Case Reset Singleton", () => TestCase.loggerResetSingleton(workbook, MSG_CELL), indent)
    run.exec("Test Case Logger: Level OFF", () => TestCase.loggerLevelOFF(), indent)
    run.exec("Test Case Logger: Counters", () => TestCase.loggerCounters(), indent)
    run.exec("Test Case Logger: Export State", () => TestCase.loggerExportState(), indent)
    run.exec("Test Case Internal Errors", () => TestCase.loggerScriptError(workbook, MSG_CELL), indent)
    /*
    run.exec("Test Case ScriptError", () => TestCase.scriptError(), indent)
    run.exec("Test Case LayoutImpl", () => TestCase.testLayoutImpl(), indent)
    run.exec("Test Case LogEventImpl", () => TestCase.testLogEventImpl(), indent)
    run.exec("Test Case ConsoleAppender", () => TestCase.consoleAppender(), indent)
    run.exec("Test Case ExcelAppender", () => TestCase.excelAppender(workbook, MSG_CELL), indent)
    run.exec("Test Case Logger: Lazy Init", () => TestCase.loggerLazyInit(), indent)
    run.exec("Test Case Reset Singleton", () => TestCase.loggerResetSingleton(workbook, MSG_CELL), indent)
    run.exec("Test Case Logger: Level OFF", () => TestCase.loggerLevelOFF(), indent)
    run.exec("Test Case Logger: Counters", () => TestCase.loggerCounters(), indent)
    run.exec("Test Case Logger: Export State", () => TestCase.loggerExportState(), indent)
    run.exec("Test Case Internal Errors", () => TestCase.loggerScriptError(workbook, MSG_CELL), indent)
    run.exec("Test Case Logger toString", () => TestCase.loggerToString(), indent)
    */

    success = true
  } catch (e) {
    // TypeScript strict mode: 'e' is of type 'unknown', so we must check its type before property access
    let info:string
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

  /** Removes the timestamp from a string. This is used to compare strings */
  public static removeTimestamp(str: string): string { // Remove timestamp from a string
    let timestampRegex = /^\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2},\d{3}\] /
    return str.replace(timestampRegex, '')
  }

  public static setShortLayout(): void { // Set the short layout for the logger
    let layout = new LayoutImpl(LayoutImpl.shortFormatterFun)
    AbstractAppender.setLayout(layout) // Set the layout for all appenders
  }

  public static setStandardLayout(): void { // Set the standard layout for the logger
    let layout = new LayoutImpl()
    AbstractAppender.setLayout(layout) // Set the layout for all appenders
  }

  /**
   * Returns a new array with only the 'type' and 'message' properties
   * from each LogEvent in the input array.
   * @param logEvents Array of LogEvent objects
   * @returns Array of objects containing only type and message
   */
  static simplifyLogEvents(
    logEvents: LogEvent[]
  ): { type: LOG_EVENT; message: string }[] {
    return logEvents.map(event => ({
      type: event.type,
      message: event.message
    }))
  }

  public static scriptError(): void { // Unit tests for the ScriptError class
    TestCase.clear() // Clear all the instances

    // Defining the variables to be used in the tests
    let expected: string, actual: string, cause: Error, origin: ScriptError

    // Testing raising a ScriptError without cause
    expected = "Script Error message"
    Assert.throws(
      () => { throw new ScriptError(expected) },
      ScriptError,
      expected,
      "scriptError(notcause)"
    )
    // Testing raising a ScriptError with cause
    cause = new TypeError("Type Error message")
    origin = new ScriptError(expected, cause)
    expected = "Script Error message (caused by 'TypeError' with message 'Type Error message')"
    Assert.throws(
      () => { throw origin },
      ScriptError,
      expected,
      "scriptError(with cause)"
    )
    // Testing re-throw
    expected = cause.message
    Assert.throws(
      () => origin.rethrowCauseIfNeeded(),
      TypeError,
      expected,
      "scriptError(rethrowIfNeeded)"
    )
    // Testing toString
    function escapeRegex(str: string): string {// Scaping metacharacters
      return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    }

    function buildRegex(trigger: ScriptError): RegExp {// Building regex for toString
      let NAME = trigger.cause ? trigger.cause.name : trigger.name
      let MSG = trigger.cause ? trigger.cause.message : trigger.message
      const regex = new RegExp(
        `^${escapeRegex(trigger.name)}: ${escapeRegex(trigger.message)}\\n` +  // Header
        `Stack trace:\\n` +                                                    // Stack section
        `${escapeRegex(NAME)}: ${escapeRegex(MSG)}\\n` +                       // type and message
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

  }

  public static testLayoutImpl(): void { // Unit tests for LayoutImpl class
    TestCase.clear()
    // Deffining the variables to be used in the tests
    let layout:Layout, event:LogEvent, actualStr:string, expectedStr:string

    layout = new LayoutImpl(LayoutImpl.shortFormatterFun) // with short formatter
    event = new LogEventImpl("Test message", LOG_EVENT.INFO)
    actualStr = layout.format(event); // This should NOT error!
    // Testing constructor
    Assert.isNotNull(layout, "LayoutImpl(constructor-is not null)")
    expectedStr = `[INFO] Test message`
    Assert.equals(actualStr, expectedStr, "LayoutImpl(format-short layout)")
    // Testing long layout
    AbstractAppender.clearLayout()
    layout = new LayoutImpl() // Default formatter with timestamp
    actualStr = TestCase.removeTimestamp(layout.format(event))
    Assert.equals(actualStr, expectedStr,"LayoutImpl(format-long layout)")
    // Testing toString: TODO: add more tests

    TestCase.clear()
  }

  public static testLogEventImpl(): void { // Unit tests for LogEventImpl class
    TestCase.clear()
    // Defining the variables to be used in the tests
    let actualEvent:LogEvent, expectedMsg: string, actualMsg: string, expectedStr, actualStr: string, 
      actualType: LOG_EVENT, actualtimeStamp: Date, layout: Layout

    layout = new LayoutImpl(LayoutImpl.shortFormatterFun) // Short layout
    // Testing constructor
    expectedMsg = "Test message"
    actualEvent = new LogEventImpl(expectedMsg, LOG_EVENT.INFO)
    Assert.isNotNull(actualEvent, "LogEventImpl(constructor-is not null)")
    // Testing properties
    actualMsg = (actualEvent as LogEvent).message
    actualType = (actualEvent as LogEvent).type
    actualtimeStamp = (actualEvent as LogEvent).timestamp
    Assert.isNotNull(actualtimeStamp, "LogEventImpl(get timestamp) is not null")
    Assert.isType(actualtimeStamp, Date, "LogEventImpl(get timestamp) is Date")
    Assert.equals(actualType, LOG_EVENT.INFO, "LogEventImpl(get type())")
    Assert.equals(actualMsg, expectedMsg, "LogEventImpl(get message())")
    // Testing toString
    expectedStr = `[${LOG_EVENT[actualType]}] ${expectedMsg}`
    actualStr = TestCase.removeTimestamp((actualEvent as LogEvent).toString())
    Assert.equals(actualStr, expectedStr, "LogEventImpl(toString())")

    TestCase.clear()
  }

  public static consoleAppender(): void { // Unit Tests for ConsoleAppender class
    TestCase.clear()
    TestCase.setShortLayout()

    // Defining the variables to be used in the tests
    let expectedStr: string, actualStr: string, expectedEvent: LogEvent,
      actualEvent: LogEvent|null, appender: Appender, layout:Layout, expectedNull: LogEvent|null,
      actualMsg: string, expectedMsg: string, msg: string

    // Test lazy initialization: We can't because we need and instance first

    // Testing log(LogEvent)
    expectedMsg = "Info Message"
    appender = ConsoleAppender.getInstance()
    expectedEvent = new LogEventImpl(expectedMsg,LOG_EVENT.INFO)
    appender.log(expectedEvent)
    actualEvent = appender.getLastLogEvent()
    Assert.equals(actualEvent?.type, expectedEvent?.type,
      "ConsoleAppender(getLastMsg not empty)-log(LogEvent).type")
    Assert.equals(actualEvent?.message, expectedEvent?.message,
      "ConsoleAppender(getLastMsg not empty)-log(LogEvent).message")

    // Testing log(string, LOG_EVENT)
    expectedMsg = "Trace Message"
    expectedEvent = new LogEventImpl(expectedMsg,LOG_EVENT.TRACE)
    appender.log(expectedMsg, LOG_EVENT.TRACE)
    actualEvent = appender.getLastLogEvent()
    Assert.equals(actualEvent?.type, expectedEvent?.type,
      "ConsoleAppender(getLastMsg not empty)-log(string,LOG_EVENT).type")
    Assert.equals(actualEvent?.message, expectedEvent?.message,
      "ConsoleAppender(getLastMsg not empty)-log(string,LOG_EVENT).message")

    // Now testing after clearing the instance the last message is empty again
    ConsoleAppender.clearInstance()
    appender = ConsoleAppender.getInstance()
    expectedNull = null
    actualEvent = appender.getLastLogEvent()
    Assert.equals(actualEvent, expectedNull,
      "ConsoleAppender(getLastMsg empty after clear the instance)")

    // Testing to String
    ConsoleAppender.clearInstance()
    msg = "error message in appendersToString"
    ConsoleAppender.clearInstance()
    ConsoleAppender.getInstance()
    actualEvent = new LogEventImpl(msg, LOG_EVENT.ERROR)
    appender.log(actualEvent)
    msg = AbstractAppender.getLayout().format(actualEvent as LogEvent)
    expectedStr = `ConsoleAppender: {Last log event: '${msg}'}`
    actualStr = appender.toString()
    Assert.equals(actualStr,expectedStr, "ConsoleAppender(toString)")

    // Testing throwing a ScriptError
    ConsoleAppender.clearInstance()
    appender = ConsoleAppender.getInstance()
    expectedStr = "[LogEventImpl.constructor]: LogEvent.type='-1' property is not defined in the LOG_EVENT enum."
    Assert.throws(
      () => appender.log("Info event message", -1 as LOG_EVENT),
      ScriptError,
      expectedStr,
      "ConsoleAppender(ScriptError)-log - non valid LOG_EVENT"
    )

    // Testing not instantiated singleton
    expectedStr = "In 'ConsoleAppender' class a singleton instance can't be undefined or null. Please invoke getInstance first"
    ConsoleAppender.clearInstance()
    Assert.throws(
      () => appender.toString(),
      ScriptError,
      expectedStr,
      "ConsoleAppender(ScriptError)-toString-with singleton not instantiated"
    )

    TestCase.clear()
  }

  // Unit Tests for ExcelAppener claass
  public static excelAppender(workbook: ExcelScript.Workbook, msgCell: string): void {
    TestCase.clear()
    TestCase.setShortLayout()
   
    // Defining the variables to be used in the tests
    let expectedStr: string, actualStr: string, msg:string, expectedEvent: LogEvent,
      actualEvent: LogEvent|null, appender: Appender, msgCellRng: ExcelScript.Range,
      activeSheet: ExcelScript.Worksheet

    activeSheet = workbook.getActiveWorksheet()
    msgCellRng = activeSheet.getRange(msgCell)
    const address = msgCellRng.getAddress()
    appender = ExcelAppender.getInstance(msgCellRng)

    // Testing sending log(string, LOG_EVENT)
    msg = "Info event in ExcelConsole"
    appender.log(msg, LOG_EVENT.INFO)
    actualStr = msgCellRng.getValue().toString()
    expectedEvent =  new LogEventImpl(msg, LOG_EVENT.INFO)
    expectedStr = AbstractAppender.getLayout().format(expectedEvent as LogEvent)
    Assert.equals(actualStr, expectedStr, "ExcelAppender(cell value via log(string,LOG_EVENT))")

    // Testing log(LogEvent)
    appender.log(expectedEvent)
    actualStr = msgCellRng.getValue().toString()
    expectedEvent =  new LogEventImpl(msg, LOG_EVENT.INFO)
    expectedStr = AbstractAppender.getLayout().format(expectedEvent as LogEvent)
    Assert.equals(actualStr, expectedStr, "ExcelAppender(cell value) via log(LogEvent)")
   
    // Testing the corresponding last log event
    actualEvent = appender.getLastLogEvent()
    Assert.isNotNull(actualEvent, "ExcelAppender(getLastEvent) not null")
    Assert.equals(actualEvent!.type, expectedEvent.type,
      "ExcelAppender(getLastEvent).type")
    Assert.equals(actualEvent!.message, expectedEvent.message,
      "ExcelAppender(getLastEvent).message")
      
    // Script Errors
    ExcelAppender.clearInstance() // singleton is undefined
    expectedStr = "In 'ExcelAppender' class a singleton instance can't be undefined or null. Please invoke getInstance first"
    Assert.throws(
      () => appender.log("Info message", LOG_EVENT.INFO),
      ScriptError,
      expectedStr,
      "ExcelAppender(ScriptError)-log"
    )
    //    Testing non valid input
     expectedStr = "ExcelAppender requires a valid ExcelScript.Range for input argument msgCellRng."
    Assert.throws(
      () => ExcelAppender.getInstance(null),
      ScriptError,
      expectedStr,
      "ExelAppender(ScriptError)-getInstance(Non valid msgCellRng-null)"
    )
    Assert.throws(
      () => ExcelAppender.getInstance(undefined),
      ScriptError,
      expectedStr,
      "ExelAppender(ScriptError))-getInstance - Non valid msgCellRng-undefined"
    )
    appender = ExcelAppender.getInstance(activeSheet.getRange(msgCell))
    expectedStr = "[LogEventImpl.constructor]: LogEvent.type='-1' property is not defined in the LOG_EVENT enum."
    Assert.throws(
      () => appender.log("Info event message", -1 as LOG_EVENT),
      ScriptError,
      expectedStr,
      "ExcelAppender(ScriptError)-Log non valid LOG_EVENT"
    )
    ExcelAppender.clearInstance()
    /*Mock object for ExcelScript.Range to simulate a multi-cell range in VS/TypeScript tests.
    This enables testing single-cell validation logic in environments where the real API isn't available.
    (Office Scripts allows any range; in VS strict typing and missing API require this manual mock.)*/
    const mockArrRng = { getCellCount: () => 2, setValue: () => {} }
    expectedStr = "ExcelAppender requires input argument msgCellRng represents a single Excel cell."
    Assert.throws(
      () => ExcelAppender.getInstance(mockArrRng as unknown as ExcelScript.Range),
      ScriptError,
      expectedStr,
      "ExcelAppender(ScriptError)-getInstance not a single cell"
    )

    //    Testing valid hexadecimal coloros
    ExcelAppender.clearInstance()
    expectedStr = "The input value 'null' color for 'error' event is not a valid hexadecimal color. Please enter a value that matches the following regular expression: '/^#?[0-9A-Fa-f]{6}$/'"
    Assert.throws(
      () => ExcelAppender.getInstance(msgCellRng, null as unknown as string), // don't use undefined, it is valid
      ScriptError,
      expectedStr,
      "ExcelAppender(ScriptError)-getInstance-red color undefined"
    )
    expectedStr = "The input value '' color for 'warning' event is not a valid hexadecimal color. Please enter a value that matches the following regular expression: '/^#?[0-9A-Fa-f]{6}$/'"
    Assert.throws(
      () => ExcelAppender.getInstance(msgCellRng, "000000", ""),
      ScriptError,
      expectedStr,
      "ExcelAppender(ScriptError)-getInstance-Non valid font color for warning"
    )
    expectedStr = "The input value 'xxxxxx' color for 'info' event is not a valid hexadecimal color. Please enter a value that matches the following regular expression: '/^#?[0-9A-Fa-f]{6}$/'"
    Assert.throws(
      () => ExcelAppender.getInstance(msgCellRng, "000000", "000000", "xxxxxx"),
      ScriptError,
      expectedStr,
      "ExcelAppender(ScriptError) - getInstance - Non valid font color for info"
    )
    expectedStr = "The input value '******' color for 'trace' event is not a valid hexadecimal color. Please enter a value that matches the following regular expression: '/^#?[0-9A-Fa-f]{6}$/'"
    Assert.throws(
      () => ExcelAppender.getInstance(msgCellRng, "000000", "000000", "000000", "******"),
      ScriptError,
      expectedStr,
      "ExcelAppender(ScriptError)-getInstance-Non valid font color for trace"
    )

    // Testing toString
    ExcelAppender.clearInstance()
    appender = ExcelAppender.getInstance(activeSheet.getRange(msgCell))
    msg = "Trace message in ExcelAppender toString"
    appender.log(msg, LOG_EVENT.TRACE)
    expectedStr = `ExcelAppender: {Message Range: "${address}", Event fonts: {errFont,9c0006,warnFont,ed7d31,infoFont,548235,traceFont,7f7f7f}, Last log event: "[TRACE] ${msg}"}`
    actualStr = (appender as ExcelAppender).toString()
    Assert.equals(actualStr, expectedStr, "ExcelAppender(toString)")

    TestCase.clear()
  }

  public static loggerLazyInit() { // Unit Tests on Lazy Initialization for Logger class (instance and appender)
    TestCase.clear()
    TestCase.setShortLayout()

    // Defining the variables to be used in the tests
    let layout: Layout, msg1: string, msg2: string, expectedEvent: LogEvent, logger: Logger, 
      expectedNum: number, actualNum: number

    // Testing lazy initialization of the LoggerImpl singleton
    msg1 = "Info event, in lazyInit"
    msg2 = "Error event, in lazyInit"
    // No appender was defined
    logger = LoggerImpl.getInstance(LoggerImpl.LEVEL.INFO)
    logger.info(msg1) // lazy initialization of the appender
    expectedNum = 1
    actualNum = logger.getAppenders().length ?? 0
    Assert.equals(actualNum, expectedNum, "Logger(Lazy init)-appender")

    // Lazy initialization of the singleton with default parameters (WARN,EXIT)
    LoggerImpl.clearInstance()
    Assert.isNotNull(LoggerImpl.getInstance(), "Lazy init(logger != null)")
    LoggerImpl.clearInstance()
    
    // Testing ScriptError when no appender is defined
    expectedEvent = new LogEventImpl(msg2, LOG_EVENT.ERROR)
    Assert.throws(
      () => logger.error(msg2),
      ScriptError,
      AbstractAppender.getLayout().format(expectedEvent as LogEvent),
      "Logger(Lazy init)-Logger"
    )


    TestCase.clear()
  }

  /** Unit tests when the Logger the level is LEVEL.OFF. 
   * Expected no log event will be sent to the appender and therefore in case
   * of error/warnings, no ScriptError will be thrown, sine no log event
   * was sent to the appenders. Under this scenario the Logger.ACTION value
   * doesn't have any effect.
   */
  public static loggerLevelOFF() {
    TestCase.clear()
    TestCase.setShortLayout()

    // Defining the variables to be used in the tests
    let layout: Layout, logger: Logger, expectedNum: number, actualNum: number

    layout = new LayoutImpl(LayoutImpl.shortFormatterFun) // Short layout

    // Testing OFF, CONTINUE
    const MSG = "event not sent in LevelOFF"
    expectedNum = 0
    actualNum = 0
    logger = LoggerImpl.getInstance(LoggerImpl.LEVEL.OFF, LoggerImpl.ACTION.CONTINUE)
    logger.error(MSG) // Not sent
    Assert.equals(expectedNum, actualNum, "LoggerLevelOff(OFF,CONTINUE)-error")
    logger.warn(MSG) // Not sent
    expectedNum = 0
    actualNum = logger.getCriticalEvents().length
    Assert.equals(expectedNum, actualNum, "LoggerLevelOff(OFF,CONTINUE)-warning")

    // Testing OFF, EXIT
    LoggerImpl.clearInstance()
    LoggerImpl.getInstance(LoggerImpl.LEVEL.OFF, LoggerImpl.ACTION.EXIT)
    logger.error(MSG)
    expectedNum = 0
    actualNum = logger.getCriticalEvents().length
    Assert.equals(expectedNum, actualNum, "LoggerLevelOFF(OFF,EXIT)-error")

    TestCase.clear()
  }

  /**Unit tests for Logger class checking the behaviour after the singleton was reset
   */
  public static loggerResetSingleton(workbook: ExcelScript.Workbook, msgCell: string) {
    TestCase.clear()
    TestCase.setShortLayout()

    // Defining the variables to be used in the tests
    let logger: Logger

    logger = LoggerImpl.getInstance()
    LoggerImpl.clearInstance() // Singleton is undefined
    // Now we need to invoke a method that doesn't invoke lazy initialization
    const EXPECTED_MSG = "In 'LoggerImpl' class a singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.getErrCnt(),
      ScriptError,
      EXPECTED_MSG,
      "loggerResetSingleton(getErrCnt())"
    )
    Assert.throws(
      () => logger.getWarnCnt(),
      ScriptError,
      EXPECTED_MSG,
      "loggerResetSingleton(getWarnCnt())"
    )
    Assert.throws(
      () => logger.hasErrors(),
      ScriptError,
      EXPECTED_MSG,
      "loggerResetSingleton(hasErrors())"
    )
    Assert.throws(
      () => logger.hasWarnings(),
      ScriptError,
      EXPECTED_MSG,
      "loggerResetSingleton(hasWarnings())"
    )
    Assert.throws(
      () => logger.getCriticalEvents(),
      ScriptError,
      EXPECTED_MSG,
      "loggerResetSingleton(getMessages())"
    )
    Assert.throws(
      () => logger.getAppenders(),
      ScriptError,
      EXPECTED_MSG,
      "loggerResetSingleton(getAppenders())"
    )
    Assert.throws(
      () => logger.getAction(),
      ScriptError,
      EXPECTED_MSG,
      "loggerResetSingleton(getAction())"
    )
    Assert.throws(
      () => logger.getLevel(),
      ScriptError,
      EXPECTED_MSG,
      "loggerResetSingleton(getLevel())"
    )

    TestCase.clear()
  }

  /**Unit Tests for Logger class for testing counters */
  public static loggerCounters() {
    TestCase.clear()
    TestCase.setShortLayout()

    // Defining the variables to be used in the tests
    let logger: Logger, layout: Layout, actualNum: number, expectedNum: number,
      actualEvent: LogEvent|null, expectedEvent: LogEvent, errMsg: string, warnMsg: string
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
    errMsg = "Error event in counters"
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
    warnMsg = "Warning event in counters"
    logger.warn(warnMsg)
    expectedNum = 1
    actualNum = logger.getWarnCnt()
    Assert.equals(actualNum, expectedNum, "loggerCounters(getWarnCnt=1)")
    // Testing messages
    let expectedArr = TestCase.simplifyLogEvents([new LogEventImpl(errMsg,LOG_EVENT.ERROR), 
      new LogEventImpl(warnMsg, LOG_EVENT.WARN)])
    let actualArr = TestCase.simplifyLogEvents(logger.getCriticalEvents())
    Assert.equals(actualArr, expectedArr, "loggerCounters(getMessages)")
    Assert.equals(logger.hasMessages(), true, "loggerCounters(hasMessages)")
    // Testing other events, don't affect the counters
    let msg = "No sent event"
    logger.info(msg)
    actualNum = logger.getErrCnt()
    Assert.equals(actualNum, expectedNum, "LoggerCounter(getErrCnt=1)")
    actualNum = logger.getWarnCnt()
    Assert.equals(actualNum, expectedNum, "LoggerCounter(getWarnCnt=1)")
    Assert.equals(actualArr, expectedArr, "loggerCounters(getMessages-2nd time)")

    // Clearing counts
    logger.clear()
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
  public static loggerToString() {
    TestCase.clear()
    TestCase.setShortLayout()
    // Defining the variables to be used in the tests
    let expected: string, actual: string, layout: Layout, logger: Logger

    layout = new LayoutImpl(LayoutImpl.shortFormatterFun) // Short layout
    logger = LoggerImpl.getInstance(LoggerImpl.LEVEL.INFO, // Level of verbose
      LoggerImpl.ACTION.CONTINUE)
    const MSGS = ["Error event in loggerToString", "Warning event in loggerToString"]
    logger.error(MSGS[0]) // lazy initialization of the appender
    logger.warn(MSGS[1])
    expected = `Logger: {Level: "INFO", Action: "CONTINUE", Error Count: "1", Warning Count: "1"}`
    actual = logger.toString()
    //console.log(logger.toString())
    Assert.equals(actual, expected, "loggerToString(Logger)")
    TestCase.clear()
  }

  /**Unit Tests for Logger class for method exportState */
  public static loggerExportState(): void {
    TestCase.clear()
    TestCase.setShortLayout()

    // Defining the variables to be used in the tests
    let logger: Logger, layout: Layout, expectedEvents: LogEvent[],
      messages: string[], state: { level: string, action: string, 
        errorCount: number, warningCount: number, criticalEvents: LogEvent[] }, msgs: string[]

    layout = new LayoutImpl(LayoutImpl.shortFormatterFun) // Short layout
    logger = LoggerImpl.getInstance(LoggerImpl.LEVEL.TRACE, LoggerImpl.ACTION.CONTINUE)
    msgs = ["warning event in exportState", "error event in exportState"]
    logger.trace("trace event in exportState")
    logger.info("info event in exportState")
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
  public static loggerScriptError(workbook: ExcelScript.Workbook, msgCell: string) {
    TestCase.clear()
    TestCase.setShortLayout()

    // Defining the variables to be used in the tests
    let logger: Logger, expectedMsg: string, actualMsg: string, appender: Appender,
      consoleAppender: ConsoleAppender, excelAppender: ExcelAppender, activeSheet: ExcelScript.Worksheet,
      msgCellRng: ExcelScript.Range
      
    // Testing non valid Logger.ACTION
    expectedMsg = "The input value level='-1', was not defined in Logger.LEVEL."
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
    expectedMsg = "In 'LoggerImpl' class a singleton instance can't be undefined or null. Please invoke getInstance first."
    Assert.throws(
      () => logger.getErrCnt(),
      ScriptError,
      expectedMsg,
      "loggerScriptError(getErrCnt())"
    )
    Assert.throws(
      () => logger.getWarnCnt(),
      ScriptError,
      expectedMsg,
      "loggerScriptError-(getWarnCnt())"
    )
    Assert.throws(
      () => logger.getCriticalEvents(),
      ScriptError,
      expectedMsg,
      "loggerScriptError-(getMessages())"
    )
    Assert.throws(
      () => logger.getLevel(),
      ScriptError,
      expectedMsg,
      "loggerScriptError(getLevel())"
    )
    Assert.throws(
      () => logger.getAction(),
      ScriptError,
      expectedMsg,
      "loggerScriptError(getAction())"
    )
    Assert.throws(
      () => logger.hasErrors(),
      ScriptError,
      expectedMsg,
      "loggerScriptError(hasErrors())"
    )
    Assert.throws(
      () => logger.hasWarnings(),
      ScriptError,
      expectedMsg,
      "loggerScriptError(hasWarnings())"
    )
    Assert.throws(
      () => logger.clear(),
      ScriptError,
      expectedMsg,
      "loggerScriptError(clear())"
    )
    // Testing adding/setting/removing appender with undefined/null singleton
    consoleAppender = ConsoleAppender.getInstance()
    Assert.throws(
      () => logger.getAppenders(),
      ScriptError,
      expectedMsg,
      "loggerScriptError(getAppenders())-undefined singleton"
    )
    Assert.throws(
      () => logger.addAppender(consoleAppender),
      ScriptError,
      expectedMsg,
      "loggerScriptError(addAppender())-undefined singleton"
    )
    Assert.throws(
      () => logger.removeAppender(consoleAppender),
      ScriptError,
      expectedMsg,
      "loggerScriptError(removeAppender())-undefined singleton"
    )
    Assert.throws(
      () => logger.setAppenders([consoleAppender, consoleAppender]),
      ScriptError,
      expectedMsg,
      "loggerScriptError(setAppenders(duplicated))-undefined singleton"
    )
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
    expectedMsg = "You can't add an appender that is null or undefined in the 'LoggerImpl' class."
    LoggerImpl.getInstance()

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
    expectedMsg = "Invalid input: 'appenders' must be a non-null array."
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
  
    expectedMsg = "Appender list contains null or undefined entry."
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
    expectedMsg = "Only one appender of type ConsoleAppender is allowed."
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
    expectedMsg = "Only one appender of type ConsoleAppender is allowed."
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
    expectedMsg = "Only one appender of type ExcelAppender is allowed."
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
  globalThis.main = main;
}
