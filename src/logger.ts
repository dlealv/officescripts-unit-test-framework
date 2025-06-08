// ===============================================
// Lightweight logging Framework for Office Script
// ===============================================

/**
 * Lightweight, extensible logging framework for Office Scripts, inspired by libraries like Log4j.
 * Enables structured logging via a singleton 'Logger', supporting multiple log levels ('Logger.LEVEL')
 * and pluggable output targets through the 'Appender' interface.
 * ### Built-in Appenders:
 * - 'ConsoleAppender': Logs to the console.
 * - ''ExcelAppender': Logs to a specified Excel cell. See 'ExcelAppender' for setup.
 * ### Error Handling Behavior:
 * - 'Logger.ACTION' controls the script's behavior for **error and warning** events only.
 *   - 'Logger.ACTION.CONTINUE': Logging continues; script execution is not halted.
 *   - 'Logger.ACTION.EXIT': Script terminates immediately on error or warning.
 * - Note: This action behavior is only triggered if the event is **actually logged**.
 *   If 'Logger.LEVEL' is set to 'Logger.LEVEL.OFF', no events are passed to appenders and
 *   'Logger.ACTION' has no effect.
 * ### Verbosity:
 * - Controlled via 'Logger.LEVEL', from 'Logger.LEVEL.ERROR' (least verbose) to
 *   'Logger.LEVEL.TRACE' (most verbose).
 * - Use 'Logger.LEVEL.OFF' to disable all logging (silent mode).
 * ### Testing Utilities:
 * - 'Assert': Basic equality and exception assertion helpers.
 * - 'TestRunner': Manages test execution and output formatting.
 * - 'TestCase': Stores and runs test functions via 'TestRunner.exec()'.
 * ### Example:
 * ```ts
 * Logger.getInstance(Logger.LEVEL.INFO, Logger.ACTION.CONTINUE) // If error or warnings continues and reports up to Info event
 * Logger.addAppender(ConsoleAppender.getInstance())
 * Logger.info("Script started") // Output: [INFO] Script started
 * ```
 * @remarks Designed and tested for Office Scripts runtime. Extendable with custom appenders via the 'Appender' interface.
 * @author David Leal
 * @date 2025-06-03
 * @version 1.0
 */


// Enum DEFINITIONS
// --------------------

/**
 * Enum representing log event types.
 * Each value corresponds to a the level of verbosity and is used internally to filter messages.
 * Logger.LEVEL static constants was implemented in a way to align with the order of the enum
 * LOG_EVENT, so ther *order* on how the LOG_EVENT values are defined matters.
 * Important:
 * - If new values are added, update the logic in related classes (e.g. ExcelAppender, Logger, etc.).
 * - Don't start the LOG_EVENT enum with 0, this value is reserved for Logger.LEVEL
 *      for not sending log events (Logger.LEVEL.OFF).
 *      It ensures the verbose level values are aligned with the LOG_EVENT enum.
 *      Logger.LEVEL is built based on LOG_EVENT, just adding as first value 0, i.e.
 *      Logger.OFF, therefore the verbosity respects the same order of the LOG_EVENT.
 * @remarks It was defined as an independent entity since it is used by appenders and by the Logger.
*/
enum LOG_EVENT {
  ERROR = 1, // Always starts with 1
  WARN,
  INFO,
  TRACE,
}

// INTERFACES
// --------------------

/**
 * Interface for all appenders.
 * Appenders handle log message delivery (e.g., to console, Excel, file, etc.). It determines the
 * channel the log event are sent.
 * Implementations must:
 * - Define how messages are logged ('log')
 * - Provide the last message sent ('getMsg')
 */
interface Appender {
  /** Sends the log event to the appender defined.
   * @throw ScriptError if the event doesn't belong to LOG_EVENT enum
   */
  log(message: string, event: LOG_EVENT): void

  /** Get the last event messages sent to the appender.
   * @throw ScriptError If the appender was not instantiated.
   */
  getLastMsg(): string

}

// CLASSES
// --------------------

/**
 * A custom error class for domain-specific or script-level exceptions.
 * Designed to provide clarity and structure when handling expected or controlled
 * failures in scripts (e.g., logging or validation errors). It supports error chaining
 * through an optional 'cause' parameter, preserving the original stack trace.
 * Prefer using 'ScriptError' for intentional business logic errors to distinguish them
 * from unexpected system-level failures.
 * @example
 * ```ts
 * const original = new Error("Missing field")
 * throw new ScriptError("Validation failed", original)
 * ```
 */
class ScriptError extends Error {
  /**
   * Constructs a new 'ScriptError'.
   * @param message A description of the error.
   * @param cause (Optional) The original error that caused this one. 
   *                         If provided the exception message will have a refernece to the cause
   */
  constructor(message: string, public cause?: Error) {
    super(message)
    this.name = new.target.name // Dinamically take the name of the class
    if (cause?.message)
      this.message += ` (caused by '${cause.constructor.name}' with message '${cause.message}')`
  }

  /**
   * Utility method to rethrow the original cause if present,
   * otherwise rethrows this 'ScriptError' itself.
   * Useful for deferring a controlled exception and then
   * surfacing the root cause explicitly.
   */
  public rethrowCauseIfNeeded(): never {
    if (this.cause) throw this.cause
    throw this
  }

  /** Override toString() method.
   * @return The name and the message on the first line, then 
   *          on the second line the Stack trace section name, i.e. 'Stack trace:'. 
   *          Starting on the third line the stack trace information. 
   *          If a cause was provided the stack trace will refer to the cause 
   *          otherwise to the original exception.
   */
  public toString(): string {
    const stack = this.cause?.stack ? this.cause.stack : this.stack
    return `${this.name}: ${this.message}\nStack trace:\n${stack}`
  }
}

/**
 * Abstract base class for all log appenders.
 * This class defines shared utility methods to standardize log formatting,
 * label generation, and event validation across concrete appender implementations.
 * Appenders such as 'ConsoleAppender' and 'ExcelAppender' should extend this class
 * to inherit consistent logging behavior.
 */
abstract class AbstractAppender {

  /**
   * Validates that the given event is a defined member of the 'LOG_EVENT' enum.
   * @param event - The log event to validate.
   * @throws ScriptError if the event is not part of the 'LOG_EVENT' enum.
   */
  protected static assertValidEvent(event: LOG_EVENT): void {
    if (!Object.values(LOG_EVENT).includes(event)) {
      throw new ScriptError(`The value '${event}' was not defined in the LOG_EVENT enum.`)
    }
  }

  /**
   * Formats a log message with its corresponding event label.
   * This method provides a consistent message format across all appenders.
   * It is **not intended to be overridden**.
   * @param msg - The original message.
   * @param event - The event type from 'LOG_EVENT'.
   * @returns The formatted message, e.g., '[WARN] Something happened'.
   */
  protected formatMsg(msg: string, event: LOG_EVENT): string {
    const LABEL = this.eventToLabel(event)
    return `${LABEL} ${msg}`
  }

  /**
   * Returns a standardized label for the given log event. It uses reverse mapping property of enum.
   * @param event - The event type from 'LOG_EVENT' enum.
   * @returns A string label, e.g., '[INFO]', '[ERROR]'.
   */
  private eventToLabel(event: LOG_EVENT): string {
    return `[${LOG_EVENT[event]}]`
  }
}

/**
 * Singleton appender that logs messages to the Office Script console
 * Usage:
 * - Call ConsoleAppender.getInstance() to get the appender
 * - Automatically used if no other appender is defined
 * @example:
 * ```ts
 * // Add console appender to the Logger
 * Logger.addAppender(ConsoleAppender.getInstance())
 * ```
*/
class ConsoleAppender extends AbstractAppender implements Appender {
  private static _instance: ConsoleAppender | null // Instance of the singleton pattern
  private _lastMsg = "" // The last log message sent by the appender
  private constructor() {
    super()
  }

  /**Process a log event according to the event type. If the singleton was no instantiated,
   * then it does lazy initialization.
   */
  public log(msg: string, event: LOG_EVENT): void {
    AbstractAppender.assertValidEvent(event)
    this.initIfNeeded() // lazy initialization
    const formatted = super.formatMsg(msg, event) // from parent
    this._lastMsg = formatted
    console.log(this._lastMsg)
  }

  /** @return the last event message sent by the appender, If it is undefined or
   * null, returns the empty string (defensive programming)
   * @throw ScriptError If the singleton was not instantiated
  */
  public getLastMsg(): string { // Get the last message sent
    ConsoleAppender.validateInstance()
    return this._lastMsg ? this._lastMsg.toString() : ""
  }

  /**Override toString method. Show the last message event sent
   * @throw ScriptError If the singleton was not instantiated
  */
  public toString(): string {
    ConsoleAppender.validateInstance()
    const name = this.constructor.name
    const MSG = this.getLastMsg()
    return (`${name}: {Last event message: '${MSG}'}`)
  }

  /**@return The singleton instance of the class. It uses lazy initialization,
   * i.e. if the singleton was not instanciated, it does a lazy initialization.
  */
  public static getInstance(): ConsoleAppender {
    if (!ConsoleAppender._instance) {
      ConsoleAppender._instance = new ConsoleAppender()
    }
    return ConsoleAppender._instance
  }

  /** Sets to null the singleton instance, usefull for running different scenarios.
   * Warning: Mainly intended for testing purposes. The state of the singleton will be lost.
   * @example There is no way to empty the last message sent after the instance was created unless
   * you reset it.
   * ```ts
   * appender:ConsoleAppender = ConsoleAppender.getInstance()
   * appender = ConsoleAppender.info("info event", LOG_EVENT.INFO)
   * appender.getLastMsg() // Output: "info event"
   * appender.clearInstance() // clear the singleton
   * appender = getInstance() // restart the singleton
   * appender.getLastMsg() // Output: ""
   * ```
  */
  public static clearInstance(): void {
    if (ConsoleAppender._instance) {
      ConsoleAppender._instance = null // Force full re-init
    }
  }

  /** Override the method from AbstractAppender. Instantiate the singleton,
   * if it was not instantiated */

  private initIfNeeded(): void {
    if (!ConsoleAppender._instance) {
      ConsoleAppender._instance = new ConsoleAppender()
    }
  }

  /** @internal
   * Common safeguard method, where calling initIfNeeded doesn't make sense.
   * @throw ScriptError In case the singleton was not initialaized
   */
  private static validateInstance() {
    if (!ConsoleAppender._instance) {
      const MSG = `In '${ConsoleAppender.name}' class a singleton instance can't be undefined or null. Please invoke getInstance first`
      throw new ScriptError(MSG)
    }
  }
}

/**
 * Singleton appender that logs messages to a specified Excel cell.
 * Logs messages in color based on the LOG_EVENT enum:
 * - ERROR: red, WARN: orange, INFO: green, TRACE: gray (defaults can be customized)
 * Usage:
 * - Must call ExcelAppender.getInstance(range) once with a valid single cell range
 * - range is used to display log messages
 * @example:
 * ```ts
 * const range = workbook.getWorksheet("Log").getRange("A1")
 * Logger.addAppender(ExcelAppender.getInstance(range)) // Add appender to the Logger
 * ```
*/
class ExcelAppender extends AbstractAppender implements Appender {
  private static readonly _COLOR_MAP = Object.freeze({ // default colors
    [LOG_EVENT.ERROR]: "9c0006",  // RED
    [LOG_EVENT.WARN]: "ed7d31",   // ORANGE
    [LOG_EVENT.INFO]: "548235",   // GREEN
    [LOG_EVENT.TRACE]: "7f7f7f"   // GRAY
  } as const)
  /* Regular expression to validate hexadecimal colors*/
  private static readonly HEX_REGEX = Object.freeze(/^#?[0-9A-Fa-f]{6}$/)

  private readonly _errFont: string = ExcelAppender._COLOR_MAP[LOG_EVENT.ERROR]
  private readonly _warnFont: string = ExcelAppender._COLOR_MAP[LOG_EVENT.WARN]
  private readonly _infoFont: string = ExcelAppender._COLOR_MAP[LOG_EVENT.INFO]
  private readonly _traceFont: string = ExcelAppender._COLOR_MAP[LOG_EVENT.TRACE]

  private static _instance: ExcelAppender | null // Instance of the singleton pattern
  private readonly _msgCellRng: ExcelScript.Range | undefined
  /*Have the last message content sent decoupled from _msgCellRng, to avoid issues with
  Excel not flushing the data on time*/
  private _lastMsg: string = ""

  // Private constructor to prevent user invocation
  private constructor(msgCellRng: ExcelScript.Range | undefined = undefined,
    errFont: string = ExcelAppender._COLOR_MAP[LOG_EVENT.ERROR],
    warnFont: string = ExcelAppender._COLOR_MAP[LOG_EVENT.WARN],
    infoFont: string = ExcelAppender._COLOR_MAP[LOG_EVENT.INFO],
    traceFont: string = ExcelAppender._COLOR_MAP[LOG_EVENT.TRACE]) {
    super()
    this._errFont = errFont
    this._warnFont = warnFont
    this._infoFont = infoFont
    this._traceFont = traceFont
    // Only call methods on _msgCellRng if it is defined
    if (this._msgCellRng) {
      this._msgCellRng.clear(ExcelScript.ClearApplyTo.contents)
      this._msgCellRng.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    }

  }

  /**
   * Returns the singleton ExcelAppender instance, creating it if it doesn't exist.
   * On first call, requires a valid single cell Excel range to display log messages and optional
   * color customizations for different log events (LOG_EVENT). Subsequent calls ignore parameters
   * and return the existing instance.
   * @param msgCellRng - Excel range where log messages will be written. Must be a single cell and
   * not null of undefined.
   * @param errFont - Hex color code for error messages (default: "9c0006" red).
   * @param warnFont - Hex color code for warnings (default: "ed7d31" orange).
   * @param infoFont - Hex color code for info messages (default: "548235" green).
   * @param traceFont - Hex color code for trace messages (default: "7f7f7f" gray).
   * @returns The singleton ExcelAppender instance.
   * @throw ScriptError if msgCellRng was not defined or if the range covers multiple cells
   *                    or if it is not a valid Excel range.
   *                    if the font colors are not valid hexadecimal values for colors
   * @example
   * ```ts
   * const range = workbook.getWorksheet("Log").getRange("A1")
   * const excelAppender = ExcelAppender.getInstance(range)
   * ExcelAppender.getInstance(range, "ff0000") // ignored if called again
   * ```
  */
  public static getInstance(
    msgCellRng: ExcelScript.Range,
    errFont: string = ExcelAppender._COLOR_MAP[LOG_EVENT.ERROR],
    warnFont: string = ExcelAppender._COLOR_MAP[LOG_EVENT.WARN],
    infoFont: string = ExcelAppender._COLOR_MAP[LOG_EVENT.INFO],
    traceFont: string = ExcelAppender._COLOR_MAP[LOG_EVENT.TRACE]): ExcelAppender {
    if (!ExcelAppender._instance) {
      if (!msgCellRng || !msgCellRng.setValue) { // Check for a valid range
        const MSG = `${ExcelAppender.name} requires a valid ExcelScript.Range for input argument msgCellRng.`
        throw new ScriptError(MSG)
      }
      if (msgCellRng.getCellCount() != 1) {
        const MSG = `${ExcelAppender.name} requires input argument msgCellRng represents a single Excel cell.`
        throw new ScriptError(MSG)
      }
      // Checking valid hexadecimal color
      ExcelAppender.assertColor(errFont, "error")
      ExcelAppender.assertColor(warnFont, "warning")
      ExcelAppender.assertColor(infoFont, "info")
      ExcelAppender.assertColor(traceFont, "trace")
      ExcelAppender._instance = new ExcelAppender(msgCellRng, errFont, warnFont, infoFont, traceFont)
    }
    // Note: No need to capture for invalid range, such as excel cel "C-1", because the input argument
    // is a range, so the error happens before the getInstance invokation, i.e. when calling getRange.

    return ExcelAppender._instance
  }

  /** Sets to null the singleton instance, usefull for running different scenarios.
  * Warning: Mainly intended for testing purposes. The state of the singleton will be lost.
  * @example
  * ```ts
  * const activeSheet = workbook.getActiveWorksheet() // workbook is input argument of main
  * const msgCellRng = activeSheet.getRange("C2")
  * appender = ExcelAppender.getInstance(msgCellRng) // with default log event colors
  * appender.info("info event") // Output: In Excel in cell C2 with green color shows: "info event"
  * // Now we want to test how getInstance() can throw a ScriptError,
  * // but we can't because the instance was already created and it is a singleton we need clearInstance
  * appender.clearInstance()
  * appender = ExcelAppender.getInstance(null) // throws a ScriptError
  * ```
  */
  public static clearInstance(): void {
    if (ExcelAppender._instance) {
      ExcelAppender._instance = null
    }
  }

  // Getters
  /** @returns Returns the last log message sent to the appender. If it is
   * undefined or null, returns the empty string (defensive programming)
   * @throw ScriptError in case the singleton was not initialized
  */
  public getLastMsg(): string {
    ExcelAppender.validateInstance()
    return this._lastMsg ? this._lastMsg.toString() : ""
  }

  /**
   * Sets the value of the cell, with the event message, using the font defined for the event type,
   * if not font was defined it doesn't change the font of the cell.
   * @param event a value from enum LOG_EVENT.
   * @throw ScriptError in case event is not a valid LOG_EVENT enun value.
   */
  public log(msg: string, event: LOG_EVENT): void {
    ExcelAppender.validateInstance();
    AbstractAppender.assertValidEvent(event);

    if (this._msgCellRng) {
      this._msgCellRng.clear(ExcelScript.ClearApplyTo.contents); // Clear the previous message (if needed)
      const FONT = ExcelAppender._COLOR_MAP[event] ?? null;
      if (FONT) {
        this._msgCellRng.getFormat().getFont().setColor(FONT);
      }
      const MSG = super.formatMsg(msg, event); // from parent
      this._msgCellRng.setValue(MSG);
      this._msgCellRng.getValue(); // Explicitly access the cell to ensure it commits the update
      this._lastMsg = MSG;
    } else { // TODO: Include in the unit testing
      const MSG = `msgCellRng in '${this.constructor.name}' class is null or undefined, please defined when calling getInstance method`
      throw new ScriptError(MSG)
    }
  }

  /**Shows instance configuration plus last message sent by the appender
  * @throw ScriptError, if the singleton was not instantiated.
 */
  public toString(): string { // Override the toString method
    ExcelAppender.validateInstance()
    const NAME = this.constructor.name
    const MSG_CELL_RNG = ExcelAppender._instance!._msgCellRng!.getAddress()
    const VALUE = this._msgCellRng!.getValue()
    const MSG = VALUE != null ? VALUE.toString() : ""
    const ERR_FONT = ExcelAppender._instance!._errFont
    const WARN_FONT = this._warnFont
    const INFO_FONT = this._infoFont
    const TRACE_FONT = this._traceFont
    return `${NAME}: {Message Range: "${MSG_CELL_RNG}", Error Font: "${ERR_FONT}", Warning Font: "${WARN_FONT}", Info Font: "${INFO_FONT}", Trace Font: "${TRACE_FONT}", Last event message: "${MSG}"}`
  }

  // Common safeguard method, where calling initIfNeeded doesn't make sense
  private static validateInstance(): void {
    if (!ExcelAppender._instance) {
      const MSG = `In '${ExcelAppender.name}' class a singleton instance can't be undefined or null. Please invoke getInstance first`
      throw new ScriptError(MSG)
    }
  }

  // Validate color is a valid exadecimal color
  private static assertColor(color: string, name: string): void {
    const match = ExcelAppender.HEX_REGEX.exec(color)
    if (match == null) {
      const MSG = `The input value '${color}' color for '${name}' event is not a valid hexadecimal color. Please enter a value that matches the following regular expression: '${ExcelAppender.HEX_REGEX.toString()}'`
      throw new ScriptError(MSG)
    }
  }

}

/**
 * Singleton class that manages application logging through appenders.
 * Supports the following log events: ERROR, WARN, INFO, TRACE (LOG_EVENT enum)
 * Supports the level of information (verbose) to show via Logger.LEVEL: OFF, ERROR, WARN, INFO, TRACE
 * If the level of information (LEVEL) if OFF, no log events will the sent to the appenders.
 * Supports the action to take in case of ERROR, WARN log events, the script can
 * continue ('Logger.ACTION.CONTINUE'), or aborts ('Logger.ACTION.EXIT'). Such actions only take effect
 * in case the LEVEL is not Logger.LEVEL.OFF.
 * Allows defining appenders, controlling the channels the events are sent.
 * It collects error/warning sent by the appenders via getMessages().
 * Usage:
 * - Initialize with Logger.getInstance(level, action)
 * - Add one or more appenders (e.g. ConsoleAppender, ExcelAppender)
 * - Use Logger.error(), warn(), info(), or trace() to log
 * Features:
 * - If no appender is added, ConsoleAppender is used by default
 * - Logs are routed through all registered appenders
 * - Collects a summary of error/warning messages and counts
 */
class Logger {
  // Constants
  public static readonly ACTION = Object.freeze({
    CONTINUE: 0, // In case of error/warning log events, the script continues
    EXIT: 1,     // In case of error/warning log event, throws an ScriptError error
  } as const)

  /*Generates the same sequence as LOG_EVENT, but adding the zero case with OFF. It ensures the numeric values
  match the values of LOG_EVENT. Note: enum can't be defined inside a class */
  public static readonly LEVEL = Object.freeze(Object.assign({ OFF: 0 }, LOG_EVENT))

  // Equivalent labels from LEVEL
  private static readonly LEVEL_LABELS = Object.entries(Logger.LEVEL).reduce((acc, [key, value]) => {
    acc[value] = key;
    return acc;
  }, {} as Record<string, string>)
  // Equivalent labels from ACION
  private static readonly ACTION_LABELS = Object.entries(Logger.ACTION).reduce((acc, [key, value]) => {
    acc[value] = key;
    return acc;
  }, {} as Record<string, string>)


  // Attributes
  private static _instance: Logger | null // Instance of the singleton pattern
  private static readonly DEFAUL_LEVEL = Logger.LEVEL.WARN
  private static readonly DEFAULT_ACTION = Logger.ACTION.EXIT
  private readonly _level: typeof Logger.LEVEL[keyof typeof Logger.LEVEL] = Logger.DEFAUL_LEVEL
  private readonly _action: typeof Logger.ACTION[keyof typeof Logger.ACTION] = Logger.DEFAULT_ACTION
  private _messages: string[] = [] // Collect all ERROR and WARN message only
  private _errCnt = 0   // Counter the number of error event found
  private _warnCnt = 0  // Counts the number of warning event found
  private _appenders: Appender[] = [] // List of appenders

  private constructor( // Private constructor to prevent user invocation
    level: typeof Logger.LEVEL[keyof typeof Logger.LEVEL] = Logger.DEFAULT_ACTION,
    action: typeof Logger.ACTION[keyof typeof Logger.ACTION] = Logger.DEFAULT_ACTION) {
    this._action = action
    this._level = level
  }

  // Getters
  /** @return A string array with error and warning event message only.
   * @throw ScriptError: If the singleton was not instantiated */
  public getMessages(): string[] {
    Logger.validateInstance()
    return Logger._instance!._messages
  }

  /** @return Total number of error message events sent to the appender
   * @throw ScriptError: If the singleton was not instantiated */
  public getErrCnt(): number {
    Logger.validateInstance()
    return Logger._instance!._errCnt
  }

  /** @return Total number of warning events sent to the appender
   * @throw ScriptError: If the singleton was not instantiated */
  public getWarnCnt(): number {
    Logger.validateInstance()
    return Logger._instance!._warnCnt
  }

  /** @return the action to take in case of errors or warning log events.
   * @throw ScriptError: If the singleton was not instantiated */
  public getAction(): typeof Logger.ACTION[keyof typeof Logger.ACTION] {
    Logger.validateInstance()
    return Logger._instance!._action
  }

  /** @throw Retuns the level of verbosity allowed in the Logger. The levels are incremental, i.e.
   * it includes all previous levels. For example: Logger.WARN includes warning and errors since
   * Logger.ERROR is lower.
   * @throw ScriptError: If the singleton was not instantiated */
  public getLevel(): typeof Logger.LEVEL[keyof typeof Logger.LEVEL] {
    Logger.validateInstance()
    return Logger._instance!._level
  }
  /**@return: Array with appenders subscribed to the Logger
   * @throw ScriptError: If the singleton was not instantiated */
  public getAppenders(): Appender[] {
    Logger.validateInstance()
    return Logger._instance!._appenders
  }

  // Setters
  /** Set the array of appenders with the input argument appenders
   * @param appender - Array with all appenders to set.
   * @throw ScriptError: - If the singleton was not instantiated,
   *                     - If the appenders is null or undefined, or containes
   *                       null or undefined entries
   *                     - If the appenders to add are not unique
   *                       by appender class. See JSDOC from addAppender.
   * @see addAppender
   */
  public setAppenders(appenders: Appender[]) {
    Logger.validateInstance()
    Logger.assertUniqueAppenderTypes(appenders)
    Logger._instance!._appenders = appenders
  }

  /**
   * @param appender - Add appender to the list of appender.
   * @throw ScriptError:
   *  - If the singleton was not instantiated.
   *  - If the input argument is null of undefined
   *  - If it breaks the class uniqueness of the appenders, i.e. all appenders must be from
   *    a different implementation of the Appender class. Appenders represent a channel
   *    for sending messsages, it doesn't make sense to send duplicated messages.
   * @see setAppenders
   */
  public addAppender(appender: Appender): void {
    Logger.validateInstance()
    if (!appender) { // It must be a valid appender
      const MSG = `You can't add an appender that is null of undefined in the '${Logger.name}' class`
      throw new ScriptError(MSG)
    }
    const newAppenders = [...Logger._instance!._appenders, appender]
    Logger.assertUniqueAppenderTypes(newAppenders)
    Logger._instance!._appenders.push(appender)
  }

  /**
   * Returns the singleton Logger instance, creating it if it doesn't exist.
   * If the Logger is created during this call, the provided 'level' and 'action'
   * parameters initialize the log level and error-handling behavior.
   * Subsequent calls ignore these parameters and return the existing instance.
   * @param level - Initial log level (default: Logger.LEVEL.WARN). Controls verbosity.
   *                Sends events to the appenders up to level of verbosity
   *                defined only. The level of verbosity is incremental, except for value
   *                Logger.LEVEL.OFF, which suppress all messages send to the appenders.
   *                For example: Logger.LEVEL.INFO, allows to send errors, warnings, and information events,
   *                but excludes trace events.
   * @param action - Action on error/warning (default: Logger.ACTION.EXIT).
   *                 Determines if the script should continue or abort.
   *                 If the value is Logger.ACTION.EXIT throws a ScriptError exception,
   *                 i.e. aborts the Script. If the action is Logger.ACTION.CONTINUE, the
   *                 scripts continues.
   * @returns The singleton Logger instance.
   * @throw ScriptError if the level input value was not defined in Logger.LEVEL.
   * @example
   * ```ts
   * // Initialize logger at INFO level, continue on errors/warnings
   * const logger = Logger.getInstance(Logger.LEVEL.INFO, Logger.ACTION.CONTINUE)
   * // Subsequent calls ignore parameters, return the same instance
   * const sameLogger = Logger.getInstance(Logger.LEVEL.ERROR, Logger.ACTION.EXIT)
   * Logger.info("Starting the Script") // send this message to all appenders
   * Logger.trace("Step one") // Doesn't sent because of Logger.LEVEL value: INFO
   * ```
   */
  public static getInstance(level: typeof Logger.LEVEL[keyof typeof Logger.LEVEL] = Logger.DEFAUL_LEVEL,
    action: typeof Logger.ACTION[keyof typeof Logger.ACTION] = Logger.DEFAULT_ACTION) {
    if (!Logger._instance) {
      Logger.assertValidLevel(level)  // Checking input value
      Logger._instance = new Logger(level, action)
    }
    return Logger._instance
  }

  /** Sets to null the singleton instance, useful for running different scenarios.
    * Warning: Mainly intended for testing purposes. The state of the singleton will be lost.
    @example
    ```ts
    // Testing how should work the logger with default configuration, and then changing the configuration.
    // Since the class, doesn't define setters methods to change the configuration, you can use
    // clearInstance to reset the singleton and instantiate it with different configuration.
    // Testing default configuration
    Logger.getInstance() //LEVEL: WARN, ACTION:EXIT
    logger.error ("error event") // Output: "error event" and ScriptError
    // Now we want to test with the following configuraiton: Logger.LEVEL:WARN, Logger.ACTION:CONTINUE
    Logger.clearInstance() // Clear the singleton
    Logger.getInstance(LEVEL.WARN,ACTION.CONTINUE)
    Logger.error("error event") // Output: "error event" (no ScriptError was thrown)
    ```ts
   */
  public static clearInstance(): void {
    if (Logger._instance) {
      Logger._instance = null // Force full re-init
    }
  }

  /**If the list of appenders is not empty, removes the appender from the list
   * @throw ScriptError: If the singleton was not instantiated
  */
  public removeAppender(appender: Appender): void {
    Logger.validateInstance()
    const appenders = Logger._instance!._appenders
    if (!Logger.isEmptyArray(appenders)) {
      const index = Logger._instance!._appenders.indexOf(appender)
      if (index > -1) { Logger._instance!._appenders.splice(index, 1) } // 1 means to delete only one element
    }
  }

  /**Sends an error log message to all appenders. If the level allows it. The level has to be greater or equal
   * than Logger.LEVEL.ERROR to send this event to the appenders. After the message is sent is updates the
   * error counter.
   * @remarks
   * If no singleton was defined it does lazy initialization with default configuration.
   * If no appender was defined, it does lazy initialization to ConsoleAppender
   * @throw ScriptError: Only if level is greater than Logger.LEVEL.OFF and the action is
   * is Logger.ACTION.EXIT
  */
  public error(msg: string): void {
    this.log(msg, LOG_EVENT.ERROR)
  }

  /**Sends warning event message to the appender. If the level allows it. The level has to be greater or equal
   * than to Logger.LEVEL.WARN to send this event to the appenders. After the message is sent, it updates the
   * warning counter.
   * @remarks
   * If no singleton was defined it does lazy initialization with default configuration.
   * If no appender was defined, it does lazy initialization to ConsoleAppender
   * @throw ScriptError: Only if level (see getInstance) is greater than Logger.LEVEL.ERROR and the action is
   * is Logger.ACTION.EXIT
   * */
  public warn(msg: string): void {
    this.log(msg, LOG_EVENT.WARN)
  }

  /**Sends info events message to the appender. If the level allows it. The level has to be greater or
   * equal to Logger.LEVEL.INFO to send this event to the appender.
   * @remarks
   * If no singleton was defined it does lazy initialization with default configuration.
   * If no appender was defined, it does lazy initialization to ConsoleAppender */
  public info(msg: string): void {
    this.log(msg, LOG_EVENT.INFO)
  }

  /**Sends trace events message to the appender. If the level allows it. The level has to be greater or
   * equal to Logger.LEVEL.TRACE to send this event to the appender.
   * @remarks
   * If no singleton was defined it does lazy initialization with default configuration.
   * If no appender was defined, it does lazy initialization to ConsoleAppender */
  public trace(msg: string): void {
    this.log(msg, LOG_EVENT.TRACE)
  }

  /**
   * @return true if an error log event was sent to the appenders, otherwise false.
   * @throw ScriptError: If the singleton was not instantiated */
  public hasErrors(): boolean {
    Logger.validateInstance()
    return Logger._instance!._errCnt > 0
  }

  /**@return true if an error log event was sent to the appenders, otherwise false.
   * @throw ScriptError: If the singleton was not instantiated */
  public hasWarnings(): boolean {
    Logger.validateInstance()
    return Logger._instance!._warnCnt > 0
  }

  /**@return true if some error or warning event has been sent by the appenders, otherwise false.
   * @throw ScriptError: If the singleton was not instantiated */
  public hasMessages(): boolean {
    Logger.validateInstance()
    return Logger._instance!._messages.length > 0
  }

  /** Resets the Logger history, i.e. state (errors, warnings, message summary). It doesn't reset the appenders.
   * @throw ScriptError: If the singleton was not instantiated */
  public clear(): void {
    Logger.validateInstance()
    Logger._instance!._messages = []
    Logger._instance!._errCnt = 0
    Logger._instance!._warnCnt = 0
  }

  /**Serializes the current state of the logger to a plain object, useful for
   * capturing logs and metrics for post-run analysis.
   * Testing/debugging: Compare expected vs actual logger state.
   * Persisting logs into Excel, JSON, or another external system.
   * @throws ScriptError If the singleton was not instantiated.
   * @return a structure with key information about the logger, such as:
   *        level, action, errorCount, warnCount, messages.
  */
  public exportState(): {
    level: string,
    action: string,
    errorCount: number,
    warningCount: number,
    messages: string[]
  } {
    Logger.validateInstance()
    const levelKey = Object.keys(Logger.LEVEL).find(k => Logger.LEVEL[k as keyof typeof Logger.LEVEL] === Logger._instance!._level);
    const actionKey = Object.keys(Logger.ACTION).find(k => Logger.ACTION[k as keyof typeof Logger.ACTION] === Logger._instance!._action);

    return {
      level: levelKey ?? "UNKNOWN",
      action: actionKey ?? "UNKNOWN",
      errorCount: Logger._instance!._errCnt,
      warningCount: Logger._instance!._warnCnt,
      messages: [...Logger._instance!._messages]
    }
  }

  /** Override toString method.
   * @throw ScriptError: If the singleton was not instantiated
   */
  public toString(): string {
    Logger.validateInstance()
    const NAME = this.constructor.name // if static method: Logger.name
    const levelTk = Object.keys(Logger.LEVEL).find(key =>
      Logger.LEVEL[key as keyof typeof Logger.LEVEL] === this._level)
    const actionTk = Object.keys(Logger.ACTION).find(key =>
      Logger.ACTION[key as keyof typeof Logger.ACTION] === this._action)
    return `${NAME}: {Level: "${levelTk}", Action: "${actionTk}", Error Count: "${Logger._instance!._errCnt}", Warning Count: "${Logger._instance!._warnCnt}"}`
  }

  /**
   * Routes the given log event message to all registred appenders.
   * Behavior:
   * - If the singleton was not instantiated, it instantiates it with default configuration (lazy initialization)
   *   it avoids sending unnecesary errors.
   * - If no appenders are defined, a 'ConsoleAppender' is automatically created and added.
   * - The message is only dispatched if the current log level allows it.
   * - If the event is of type 'ERROR' or 'WARN', it is recorded internally and counted.
   * - If the configured action is 'Logger.ACTION.EXIT', a 'ScriptError' is thrown for errors and warnings.
   * @throw ScriptError in case the action defined for the logger is Logger.ACTION.EXIT and the event type
   *      is LOG_EVENT.ERROR, LOG_EVENT.WARN
   */
  private log(msg: string, event: LOG_EVENT): void {
    Logger.initIfNeeded() // lazy initialization of the singleton with default parameters
    const SEND_EVENTS = Logger._instance!._level != Logger.LEVEL.OFF
    if (Logger.isEmptyArray(this.getAppenders())) {
      this.addAppender(ConsoleAppender.getInstance()) // lazy initialization at least the basic appender
    }
    if (SEND_EVENTS) { // Sends events through out the appenders
      if (Logger._instance!._level >= event) { // only if the verbose level allows it
        for (const appender of Logger._instance!._appenders) { // sends to all appenders
          appender.log(msg, event) // It can't be null/undefined at this point (no need to prevent it)
        }
      }
    }

    if (SEND_EVENTS && (event <= LOG_EVENT.WARN)) {// Only collects errors or warnings event messages
      // Updating the counter
      if (event === LOG_EVENT.ERROR) ++Logger._instance!._errCnt
      if (event === LOG_EVENT.WARN) ++Logger._instance!._warnCnt
      // Updating the message. Assumes first appender is representative (message for all apenders are the same)
      const LAST_MSG = Logger._instance!._appenders[0].getLastMsg()
      Logger._instance!._messages.push(LAST_MSG)
      if (Logger._instance!._action === Logger.ACTION.EXIT) {
        throw new ScriptError(LAST_MSG)
      }
    }
  }

  /** Returns the corresonding string label for the level. */
  private static getLevelLabel(): string {
    return Logger.LEVEL_LABELS[Logger._instance!._level]
  }

  /** Returns the corresonding string label for the action. */
  private static getActionLabel(): string {
    return Logger.ACTION_LABELS[Logger._instance!._action]
  }

  /* Enforce instantiation lazily, if the user didn't invoke getInstance(), provides a logger
   * with default configuration. It also send a trace event indicating the lazy initialization*/
  private static initIfNeeded(): void {
    if (!Logger._instance) {
      Logger._instance = Logger.getInstance()
      const LEVEL_LABEL = `Logger.LEVEL.${Logger.getLevelLabel()}`
      const ACTION_LABEL = `Logger.ACTION.${Logger.getActionLabel()}`
      const MSG = `Logger instantiated via Lazy initialization with default parameters (level=${LEVEL_LABEL}, action=${ACTION_LABEL})`
      Logger._instance.trace(MSG)
    }
  }

  // Common safeguard method, where calling initIfNeeded doesn't make sense
  private static validateInstance() {
    if (!Logger._instance) {
      const MSG = `In '${Logger.name}' class a singleton instance can't be undefined or null. Please invoke getInstance first`
      throw new ScriptError(MSG)
    }
  }

  private static isEmptyArray<T>(arr: T[]): boolean { // Helper method to check for an empty array
    return (!Array.isArray(arr) || !arr.length) ? true : false
  }

  /*Check level has one of the valid values. It is required, because the way Logger.LEVEL was built,
  i.e. based on LOG_LEVEL, so it doesn't check for non valid values during compilation. That is not the
  case of Logger.ACTION */
  private static assertValidLevel(level: typeof Logger.LEVEL[keyof typeof Logger.LEVEL]) {
    if (!Object.values(Logger.LEVEL).includes(level)) { // level not part of Logger.LEVEL
      const MSG = `The input value level='${level}', was not defined in Logger.LEVEL.`
      throw new ScriptError(MSG)
    }
  }

  /**Validates that all appenders are of unique class types, no null or undefined for
   * for appenders and all its entries.  */
  private static assertUniqueAppenderTypes(appenders: (Appender | null | undefined)[]): void {
    if (Logger.isEmptyArray(appenders)) {
      throw new ScriptError("Invalid input: 'appenders' must be a non-null array.")
    }

    const seen = new Set<Function>() // ensure unique elements only
    for (const appender of appenders) {
      if (!appender) {
        throw new ScriptError("Appender list contains null or undefined entry.")
      }
      const ctor = appender.constructor
      if (seen.has(ctor)) {
        const name = ctor.name || "UnknownAppender"
        throw new ScriptError(`Only one appender of type ${name} is allowed.`)
      }
      seen.add(ctor)
    }
  }

  /**
   * Debug-only method to reset the singleton instance.
   * Use this to simulate a fresh Logger instance during testing.
   * WARNING: Should not be used in production.
   */
  /*
  public static __debugReset(): void {
      Logger._instance = undefined as unknown as Logger
  }
  */
}

// ===================================================
// End Lightweight logging framework for Office Script
// ===================================================

// Make Logger and ConsoleAppender available globally for Node/ts-node test environments
if (typeof globalThis !== "undefined") {
  if (typeof Logger !== "undefined") {
    // @ts-ignore
    globalThis.Logger = Logger
  }
  if (typeof ConsoleAppender !== "undefined") {
    // @ts-ignore
    globalThis.ConsoleAppender = ConsoleAppender;
  }
  if (typeof ExcelAppender !== "undefined") {
    // @ts-ignore
    globalThis.ExcelAppender = ExcelAppender
  }
  if (typeof ScriptError !== "undefined") {
    // @ts-ignore
    globalThis.ScriptError = ScriptError
  }
  if (typeof LOG_EVENT !== "undefined") {
    // @ts-ignore
    globalThis.LOG_EVENT = LOG_EVENT
  }
}