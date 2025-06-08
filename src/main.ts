// ----------------------------------------
// Testing the Logging framework
// ----------------------------------------

//Main function of the Script
function main(workbook: ExcelScript.Workbook,
) {

  console.log("main.ts executed!", workbook);

  // Parameters and constants definitions
  // ------------------------------------
  const MSG_CELL = "C2" // relative to the active sheet
  //const VERBOSITY = TestRunner.VERBOSITY.OFF        // unmment the scenario of your preference
  const VERBOSITY = TestRunner.VERBOSITY.HEADER
  //const VERBOSITY = TestRunner.VERBOSITY.SECTION
  const START_TEST = "START TEST"
  const END_TEST = "END TEST"
  const SHOW_TRACE = false

  let run: TestRunner = new TestRunner(VERBOSITY) // Controles the test execution process
  let success = false // Control variable to send the last message in finally

  // MAIN EXECUTION
  // --------------------

  try {
    const VERBOSITY_LEVEL = run.getVerbosityLabel()
    run.title(`${START_TEST} with verbosity '${VERBOSITY_LEVEL}'`, 1)
    let indent: number = 2 // Use the same indentation level for all test cases
    /*All functions need to be invoked using arrow function (=>).
    Test cases organized by topics. They don't have any dependency, so they can
    be executed in any order.*/
    run.exec("Test Case ScriptError", () => TestCase.scriptError(), indent)
    run.exec("Test Case ConsoleAppender", () => TestCase.consoleAppender(), indent)
    run.exec("Test Case ExcelAppender", () => TestCase.excelAppender(workbook, MSG_CELL), indent)
    run.exec("Test Case Logger: Lazy Init", () => TestCase.loggerLazyInit(), indent)
    run.exec("Test Case Reset Singleton", () => TestCase.loggerResetSingleton(workbook, MSG_CELL), indent)
    run.exec("Test Case Logger: Level OFF", () => TestCase.loggerLevelOFF(), indent)
    run.exec("Test Case Logger: Counters", () => TestCase.loggerCounters(), indent)
    run.exec("Test Case Logger: Export State", () => TestCase.loggerExportState(), indent)
    run.exec("Test Case Internal Errors", () => TestCase.loggerScriptError(workbook, MSG_CELL), indent)
    run.exec("Test Case Logger toString", () => TestCase.loggerToString(), indent) // TODO Split

    success = true
  } catch (e) {
    // TypeScript strict mode: 'e' is of type 'unknown', so we must check its type before property access
    let info
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
    TestCase.clearAllSingleton() // safeguard
    run.title(success ? `${END_TEST}: OK` : `${END_TEST}: FAIL`, 1)
  }
} // End of main

// Testing Classes
// -----------------

/**Encapsulates the test cases to be executed as static methods of this class. To be
 * executed via TestRunner.exec method.
 */
class TestCase {

  public static clearAllSingleton(): void { // Clear all the instances
    Logger.clearInstance()
    ConsoleAppender.clearInstance()
    ExcelAppender.clearInstance()

  }

  public static scriptError(): void { // Unit tests for the ScriptError class
    // Testing raising a ScriptError without cause
    let expected = "Script Error message"
    Assert.throws(
      () => { throw new ScriptError(expected) },
      ScriptError,
      expected,
      "scriptError(notcause)"
    )
    // Testing raising a ScriptError with cause
    let cause = new TypeError("Type Error message")
    let origin = new ScriptError(expected, cause)
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

  public static consoleAppender(): void { // Unit Tests for ConsoleAppender class
    TestCase.clearAllSingleton()
    // Test lazy initialization: We can't because we need and instance first
    // if we invoke getLastMsg, it will throw a ScriptError.

    // Testing getLastMsg()
    let appender: Appender = ConsoleAppender.getInstance()
    let expected = ""
    let actual = appender.getLastMsg()
    Assert.equals(actual, expected,
      "ConsoleAppender(getLastMsg empty)")
    expected = "[INFO] Info message"
    appender.log("Info message", LOG_EVENT.INFO)
    actual = appender.getLastMsg()
    Assert.equals(actual, expected,
      "ConsoleAppender(getLastMsg not empty)")

    // Now testing after clearing the instance the last message is empty again
    ConsoleAppender.clearInstance()
    expected = ""
    appender = ConsoleAppender.getInstance()
    actual = appender.getLastMsg()
    Assert.equals(actual, expected,
      "ConsoleAppender(getLastMsg empty after clear the instance)")

    // Testing to String
    ConsoleAppender.clearInstance()
    const MSG = "error message in appendersToString"
    ConsoleAppender.clearInstance()
    ConsoleAppender.getInstance()
    appender.log(MSG, LOG_EVENT.ERROR)
    expected = `ConsoleAppender: {Last event message: '[ERROR] ${MSG}'}`
    actual = appender.toString()
    Assert.equals(actual, expected, "ConsoleAppender(toString)")

    // Testing throwing a ScriptError
    /*
    appender = ConsoleAppender.getInstance()
    expected = "The value '-1' was not defined in the LOG_EVENT enum."
    Assert.throws(
      () => appender.log("Info event message", -1),
      ScriptError,
      expected,
      "ConsoleAppender(ScriptError)-log - non valid LOG_EVENT"
    )
    */

    // Testing not instantiated singleton
    ConsoleAppender.clearInstance()
    expected = "In 'ConsoleAppender' class a singleton instance can't be undefined or null. Please invoke getInstance first"
    Assert.throws(
      () => appender.getLastMsg(),
      ScriptError,
      expected,
      "ConsoleAppender(ScriptError)-getLastMsg-singleton not instantiated"
    )
    Assert.throws(
      () => appender.toString(),
      ScriptError,
      expected,
      "ConsoleAppender(ScriptError)-toString-with singleton not instantiated"
    )

    TestCase.clearAllSingleton()
  }

  // Unit Tests for ExcelAppener claass
  public static excelAppender(workbook: ExcelScript.Workbook, msgCell: string): void {
    TestCase.clearAllSingleton()
    const activeSheet = workbook.getActiveWorksheet()
    const msgCellRng = activeSheet.getRange(msgCell)
    const address = activeSheet.getRange(msgCell).getAddress()
    let appender: ExcelAppender = ExcelAppender.getInstance(msgCellRng)

    // Testing sending log events using directly the appender
    const MSG = "Info event in ExcelConsole"
    appender.log(MSG, LOG_EVENT.INFO)
    let actual = msgCellRng.getValue().toString()
    let expected = `[INFO] ${MSG}`
    Assert.equals(actual, expected, "ExcelAppender(cell value)")
    actual = appender.getLastMsg()
    Assert.equals(actual, expected, "ExcelAppender(getLastMsg())")

    // Script Error
    appender = ExcelAppender.getInstance(msgCellRng)
    ExcelAppender.clearInstance() // singleton is undefined
    expected = "In 'ExcelAppender' class a singleton instance can't be undefined or null. Please invoke getInstance first"
    Assert.throws(
      () => (appender as ExcelAppender).getLastMsg(),
      ScriptError,
      `${expected}`,
      "ExcelAppender(ScriptError)-getLastMsg"
    )
    Assert.throws(
      () => (appender as ExcelAppender).log("Info message", LOG_EVENT.INFO),
      ScriptError,
      `${expected}`,
      "ExcelAppender(ScriptError)-log"
    )

    //    Testing non valid input
    expected = "ExcelAppender requires a valid ExcelScript.Range for input argument msgCellRng."
    /*
    Assert.throws(
      () => ExcelAppender.getInstance(null),
      ScriptError,
      expected,
      "ExelAppender(ScriptError)-getInstance(Non valid msgCellRng-null)"
    )

    Assert.throws(
      () => ExcelAppender.getInstance(undefined),
      ScriptError,
      expected,
      "ExelAppender(ScriptError))-getInstance - Non valid msgCellRng-undefined"
    )
    */
    /*
    appender = ExcelAppender.getInstance(activeSheet.getRange(msgCell))
    expected = "The value '-1' was not defined in the LOG_EVENT enum."
    Assert.throws(
      () => appender.log("Info event message", -1),
      ScriptError,
      expected,
      "ExcelAppender(ScriptError)-Log non valid LOG_EVENT"
    )
    */
    ExcelAppender.clearInstance()
    const arrayRng = activeSheet.getRange("C2:C3")
    expected = "ExcelAppender requires input argument msgCellRng represents a single Excel cell."
    Assert.throws(
      () => ExcelAppender.getInstance(arrayRng),
      ScriptError,
      expected,
      "ExcelAppender(ScriptError)-getInstance not a single cell"
    )

    //    Testing valid hexadecimal coloros
    /*
    expected = "The input value 'null' color for 'error' event is not a valid hexadecimal color. Please enter a value that matches the following regular expression: '/^#?[0-9A-Fa-f]{6}$/'"
    Assert.throws(
      () => ExcelAppender.getInstance(msgCellRng, null),
      ScriptError,
      expected,
      "ExcelAppender(ScriptError)-getInstance-Non valid font color for error"
    )
    */
    expected = "The input value '' color for 'warning' event is not a valid hexadecimal color. Please enter a value that matches the following regular expression: '/^#?[0-9A-Fa-f]{6}$/'"
    Assert.throws(
      () => ExcelAppender.getInstance(msgCellRng, "000000", ""),
      ScriptError,
      expected,
      "ExcelAppender(ScriptError)-getInstance-Non valid font color for warning"
    )
    expected = "The input value 'xxxxxx' color for 'info' event is not a valid hexadecimal color. Please enter a value that matches the following regular expression: '/^#?[0-9A-Fa-f]{6}$/'"
    Assert.throws(
      () => ExcelAppender.getInstance(msgCellRng, "000000", "000000", "xxxxxx"),
      ScriptError,
      expected,
      "ExcelAppender(ScriptError) - getInstance - Non valid font color for info"
    )
    expected = "The input value '******' color for 'trace' event is not a valid hexadecimal color. Please enter a value that matches the following regular expression: '/^#?[0-9A-Fa-f]{6}$/'"
    Assert.throws(
      () => ExcelAppender.getInstance(msgCellRng, "000000", "000000", "000000", "******"),
      ScriptError,
      expected,
      "ExcelAppender(ScriptError)-getInstance-Non valid font color for trace"
    )

    // Testing toString
    ExcelAppender.clearInstance()
    appender = ExcelAppender.getInstance(activeSheet.getRange(msgCell))
    appender.log(MSG, LOG_EVENT.ERROR)
    expected = `ExcelAppender: {Message Range: "${address}", Error Font: "9c0006", Warning Font: "ed7d31", Info Font: "548235", Trace Font: "7f7f7f", Last event message: "[ERROR] ${MSG}"}`
    actual = (appender as ExcelAppender).toString()
    Assert.equals(actual, expected, "ExcelAppender(toString)")

    TestCase.clearAllSingleton()
  }

  public static loggerLazyInit() { // Unit Tests on Lazy Initialization for Logger class (instance and appender)
    TestCase.clearAllSingleton()

    // No appender was defined
    let logger = Logger.getInstance(Logger.LEVEL.INFO)
    const MSG1 = "Info event, in lazyInit"
    const MSG2 = "Error event, in lazyInit"
    logger.info(MSG1) // lazy initialization of the appender
    const EXPECTED_NUM = 1
    const ACTUAL_NUM = logger.getAppenders().length ?? 0
    Assert.equals(ACTUAL_NUM, EXPECTED_NUM, "Logger(Lazy init)-appender")

    // Lazy initialization of the singleton with default parameters (WARN,EXIT)
    Logger.clearInstance()
    Assert.isNotNull(Logger.getInstance(), "Lazy init(logger != null)")
    Logger.clearInstance()
    Assert.throws(
      () => logger.error(MSG2),
      ScriptError,
      `[ERROR] ${MSG2}`,
      "Logger(Lazy init)-Logger"
    )

    TestCase.clearAllSingleton()
  }

  /** Unit tests when the Logger the level is LEVEL.OFF. 
   * Expected no log event will be sent to the appender and therefore in case
   * of error/warnings, no ScriptError will be thrown, sine no log event
   * was sent to the appenders. Under this scenario the Logger.ACTION value
   * doesn't have any effect.
   */
  public static loggerLevelOFF() {
    TestCase.clearAllSingleton()

    // Testing OFF, CONTINUE
    const MSG = "event not sent in LevelOFF"
    let logger = Logger.getInstance(Logger.LEVEL.OFF, Logger.ACTION.CONTINUE)
    logger.error(MSG) // Not sent
    let expectedNum = 0
    let actual = logger.getMessages().length
    Assert.equals(expectedNum, actual, "LoggerLevelOff(OFF,CONTINUE)-error")
    logger.warn(MSG) // Not sent
    expectedNum = 0
    actual = logger.getMessages().length
    Assert.equals(expectedNum, actual, "LoggerLevelOff(OFF,CONTINUE)-warning")

    // Testing OFF, EXIT
    Logger.clearInstance()
    Logger.getInstance(Logger.LEVEL.OFF, Logger.ACTION.EXIT)
    logger.error(MSG)
    expectedNum = 0
    actual = logger.getMessages().length
    Assert.equals(expectedNum, actual, "LoggerLevelOFF(OFF,EXIT)-error")
    TestCase.clearAllSingleton()
  }

  /**Unit tests for Logger class checking the behaviour after the singleton was reset
   */
  public static loggerResetSingleton(workbook: ExcelScript.Workbook, msgCell: string) {
    TestCase.clearAllSingleton()
    let logger = Logger.getInstance()
    Logger.clearInstance() // Singleton is undefined
    // Now we need to invoke a method that doesn't invoke lazy initialization
    const EXPECTED_MSG = "In 'Logger' class a singleton instance can't be undefined or null. Please invoke getInstance first"
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
      () => logger.getMessages(),
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

    TestCase.clearAllSingleton()
  }

  /**Unit Tess for Logger class for testing counters */
  public static loggerCounters() {
    TestCase.clearAllSingleton()
    let logger = Logger.getInstance(Logger.LEVEL.INFO, Logger.ACTION.CONTINUE)
    logger.addAppender(ConsoleAppender.getInstance())
    // Testing counters on initial state
    Assert.equals(logger.getErrCnt(), 0, "loggerCounters(getErrCnt=0)")
    Assert.equals(logger.getWarnCnt(), 0, "loggerCounters(getWarnCnt=0)")
    Assert.equals(logger.hasErrors(), false, "loggerCounters(hasErrors=false)")
    Assert.equals(logger.hasWarnings(), false, "loggerCounters(hasWarnings=false)")
    Assert.equals(logger.getMessages(), [], "loggerCounters(getMessages=[])")

    // Sending events affecting the counter
    const ERR_MSG = "Error event in counters"
    logger.error(ERR_MSG)
    let expectedNum = 1
    let actualNum = logger.getErrCnt()
    Assert.equals(actualNum, expectedNum, "loggerCounters(getErrCnt=1)")
    let actualStr = logger.getMessages()[0]
    let expectedStr = `[ERROR] ${ERR_MSG}`
    Assert.equals(actualStr, expectedStr, "LoggerCounters(getMessage()[0])")

    // Testing counter for warnings
    let WARN_MSG = "Warning event in counters"
    logger.warn(WARN_MSG)
    expectedNum = 1
    actualNum = logger.getWarnCnt()
    Assert.equals(actualNum, expectedNum, "loggerCounters(getWarnCnt=1)")
    // Testing messages
    let expectedArr = [`[ERROR] ${ERR_MSG}`, `[WARN] ${WARN_MSG}`]
    let actualArr = logger.getMessages()
    Assert.equals(actualArr, expectedArr, "loggerCounters(getMessages)")
    Assert.equals(logger.hasMessages(), true, "loggerCounters(hasMessages)")
    // Testing other events, don't affect the counters
    let MSG = "No sent event"
    logger.info(MSG)
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
    actualArr = logger.getMessages()
    Assert.equals(actualArr, expectedArr, "LoggerCounter(messages cleared)")
    // Checking appenders were not removed
    expectedNum = 1
    actualNum = logger.getAppenders().length
    Assert.equals(actualNum, expectedNum, "LoggerCounter(appenders not removed)")
    TestCase.clearAllSingleton()
  }

  /**Unit Tests for Logger class on toString method */
  public static loggerToString() {
    TestCase.clearAllSingleton()
    let logger = Logger.getInstance(Logger.LEVEL.INFO, // Level of verbose
      Logger.ACTION.CONTINUE)
    const MSGS = ["Error event in loggerToString", "Warning event in loggerToString"]
    logger.error(MSGS[0]) // lazy initialization of the appender
    logger.warn(MSGS[1])
    let expected = `Logger: {Level: "INFO", Action: "CONTINUE", Error Count: "1", Warning Count: "1"}`
    let actual = Logger.getInstance().toString()
    //console.log(logger.toString())
    Assert.equals(actual, expected, "loggerToString(Logger)")
    TestCase.clearAllSingleton()
  }

  /**Unit Tests for Logger class for method exportState */
  public static loggerExportState(): void {
    TestCase.clearAllSingleton()
    let logger = Logger.getInstance(Logger.LEVEL.TRACE, Logger.ACTION.CONTINUE)
    const MSGS = ["warning event in exportState", "error event in exportState"]
    logger.trace("trace event in exportState")
    logger.info("info event in exportState")
    logger.warn(MSGS[0])
    logger.error(MSGS[1])
    const state = logger.exportState()
    Assert.equals(state.level, "TRACE", "loggerExportState(level)")
    Assert.equals(state.action, "CONTINUE", "loggerExportState(action)")
    Assert.equals(state.errorCount, 1, "loggerExportState(errorCount)")
    Assert.equals(state.warningCount, 1, "loggerExportState(warningCount)")
    Assert.equals(state.messages.length, 2, "loggerExportState(messages.length)")
    const messages: string[] = [`[WARN] ${MSGS[0]}`, `[ERROR] ${MSGS[1]}`]
    Assert.equals(state.messages, messages, "loggerExportState(messages))")
    TestCase.clearAllSingleton()
  }

  /**Unit tests for Logger class, for testing scenarios where a ScriptError will be thrown.
   * It also tests all defensing programming scenarios implemented.
   */
  public static loggerScriptError(workbook: ExcelScript.Workbook, msgCell: string) {
    TestCase.clearAllSingleton()
    Logger.clearInstance()

    // Testing non valid Logger.ACTION
    let expectedMsg = "The input value level='-1', was not defined in Logger.LEVEL."
    Assert.throws(
      () => Logger.getInstance(-1, Logger.ACTION.CONTINUE),
      ScriptError,
      expectedMsg,
      "loggerScriptError-Non valid LOG_LEVEL enum value"
    )

    // Testing when is invoked validateInstance method
    Logger.clearInstance()
    let logger = Logger.getInstance()
    Logger.clearInstance() // now _instance is undefined
    const EXPECTED_MSG = "In 'Logger' class a singleton instance can't be undefined or null. Please invoke getInstance first"
    Assert.throws(
      () => logger.getErrCnt(),
      ScriptError,
      EXPECTED_MSG,
      "loggerScriptError(getErrCnt())"
    )
    Assert.throws(
      () => logger.getWarnCnt(),
      ScriptError,
      EXPECTED_MSG,
      "loggerScriptError-(getWarnCnt())"
    )
    Assert.throws(
      () => logger.getMessages(),
      ScriptError,
      EXPECTED_MSG,
      "loggerScriptError-(getMessages())"
    )
    Assert.throws(
      () => logger.getAppenders(),
      ScriptError,
      EXPECTED_MSG,
      "loggerScriptError(getAppenders())"
    )
    Assert.throws(
      () => logger.getLevel(),
      ScriptError,
      EXPECTED_MSG,
      "loggerScriptError(getLevel())"
    )
    Assert.throws(
      () => logger.getAction(),
      ScriptError,
      EXPECTED_MSG,
      "loggerScriptError(getAction())"
    )
    Assert.throws(
      () => logger.hasErrors(),
      ScriptError,
      EXPECTED_MSG,
      "loggerScriptError(hasErrors())"
    )
    Assert.throws(
      () => logger.hasWarnings(),
      ScriptError,
      EXPECTED_MSG,
      "loggerScriptError(hasWarnings())"
    )
    Assert.throws(
      () => logger.clear(),
      ScriptError,
      EXPECTED_MSG,
      "loggerScriptError(clear())"
    )
    // Testing add appenders under singleton cleared
    const activeSheet = workbook.getActiveWorksheet()
    const sheetName = activeSheet.getName()
    let consoleAppender = ConsoleAppender.getInstance()
    Assert.throws(
      () => logger.addAppender(consoleAppender),
      ScriptError,
      EXPECTED_MSG,
      "loggerScriptError(addAppender())"
    )
    Assert.throws(
      () => logger.removeAppender(consoleAppender),
      ScriptError,
      EXPECTED_MSG,
      "loggerScriptError(removeAppender())"
    )
    Assert.throws(
      () => logger.setAppenders([consoleAppender, consoleAppender]),
      ScriptError,
      EXPECTED_MSG,
      "loggerScriptError(setAppenders(duplicated))"
    )
    Assert.throws(
      () => logger.toString(),
      ScriptError,
      EXPECTED_MSG,
      "loggerScriptError(toString())"
    )
    // Testing adding a null/undefined appender
    Logger.clearInstance()
    expectedMsg = "You can't add an appender that is null of undefined in the 'Logger' class"
    Logger.getInstance()
    Assert.throws(
      () => logger.addAppender(null as unknown as Appender),
      ScriptError,
      expectedMsg,
      "loggerScriptError(addAppenders()-null)"
    )
    Assert.throws(
      () => logger.addAppender(undefined as unknown as Appender),
      ScriptError,
      expectedMsg,
      "loggerScriptError(addAppenders()-undefined)"
    )
    // Adding appenders via setAppenders
    /*
    expectedMsg = "Invalid input: 'appenders' must be a non-null array."
    Assert.throws(
      () => logger.setAppenders(undefined),
      ScriptError,
      expectedMsg,
      `Internal Error(setAppenders)-undefined`
    )
    Assert.throws(
      () => logger.setAppenders(null),
      ScriptError,
      expectedMsg,
      "loggerScriptError(setAppenders)-null"
    )
    expectedMsg = "Appender list contains null or undefined entry."
    Assert.throws(
      () => logger.setAppenders([consoleAppender, null]),
      ScriptError,
      expectedMsg,
      "loggerScriptError-[consoleAppender,null]"
    )
    Assert.throws(
      () => logger.setAppenders([consoleAppender, undefined]),
      ScriptError,
      expectedMsg,
      "loggerScriptError-[consoleAppender,undefined]"
    )
    */
    expectedMsg = "Only one appender of type ConsoleAppender is allowed."
    Assert.throws(
      () => logger.setAppenders([consoleAppender, consoleAppender]),
      ScriptError,
      expectedMsg,
      "loggerScriptError-[consoleAppender,consoleAppender]"
    )

    // Testing adding duplicate appender
    Logger.clearInstance()
    logger = Logger.getInstance()
    logger.addAppender(ConsoleAppender.getInstance())
    expectedMsg = "Only one appender of type ConsoleAppender is allowed."
    Assert.throws(
      () => logger.addAppender(ConsoleAppender.getInstance()),
      ScriptError,
      expectedMsg,
      "loggerScriptError-addaAppender duplicated"
    )
    Logger.clearInstance()
    Logger.getInstance()
    let excelAppender = ExcelAppender.getInstance(activeSheet.getRange(msgCell))
    expectedMsg = "Only one appender of type ExcelAppender is allowed."
    Assert.throws(
      () => logger.setAppenders([excelAppender, excelAppender]),
      ScriptError,
      expectedMsg,
      "loggerScriptError-setAppender - duplicated"
    )

    TestCase.clearAllSingleton()
  }

}

// Make main available globally for Node/ts-node test environments
if (typeof globalThis !== "undefined" && typeof main !== "undefined") {
  // @ts-ignore
  globalThis.main = main;
}