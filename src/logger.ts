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
 * version 1.2.0
 * date 2025-06-03
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

// Types
// --------------------

/**
 * Function type for creating a LogEvent. Used in appender implementations.
 * @param eventType - The log event type.
 * @param message - The log message.
 * @returns A LogEvent object.
 */
type LogEventFactory = (message: string,eventType: LOG_EVENT) => LogEvent


// INTERFACES
// --------------------

/**
 * Interface for all log events to be sent to the appenders.
 * It defines the structure of a log event, intended to be immutable.
 * 
 * @remarks The layout is used to configure the core content formatting of the log event message.
 * All appenders (listeners) will receive the same formatted log content, ensuring consistency.
 * Appenders may apply their own additional presentation (e.g., colors or styles) without altering 
 * the event message itself.
 * 
 * Although getLayout/setLayout are instance methods, a typical implementation shares a single static 
 * layout for all events.
 */

interface LogEvent {
  /** The event type from LOG_EVENT enum. It is immutable */
  readonly type: LOG_EVENT
  /** The log message to be sent. It is inmmutable. */
  readonly message: string
  /** The timestamp when the event was created. It is immutable */
  readonly timestamp: Date 
  /** Returns a string representation of the log event.*/
  toString(): string  
}

/**
 * Interface to handle formatting of log events sent to appenders. The formant means
 * how the content of the message is structured before sending it to the appenders. It is not intened
 * for adormentation or presentation of the log event, but rather to define the core content.
 * 
 * @remarks Implementations should provide consistent formatting for all log events.
 * While the interface does not enforce singleton/static usage, typical implementations
 * may use a shared instance to ensure consistency across all events.
 */
interface Layout {
  /** The layout used to format the log event before sending it to the appenders. Since the layout should be the same
   * for all log events it should be defined as static and initialized following the singleton pattern.
   */

  /** Formats the log event into a string representation.
   * @param event - The log event to format.
   * @returns A formatted string representation of the log event.
   */
  format(event: LogEvent): string

  /** Returns a string representation of the layout.*/
  toString(): string
}

/**
 * Interface for all appenders.
 * Appenders handle log message delivery (e.g., to console, Excel, file, etc.). It determines the
 * channel the log event are sent.
 * Implementations must:
 * - Define how messages are logged ('log')
 * - Provide the last message sent ('getLastLogEvent')
 */
interface Appender {
  /**
   * Sends the log event to the appender defined.
   * @param event The log event object.
   * @throws ScriptError if the event is invalid.
   */
  log(event: LogEvent): void

  /**
   * Sends a log message to the appender based on the event type.
   * @param msg The message to log.
   * @param event The type of log event (from LOG_EVENT enum).
   */
  log(msg: string, event: LOG_EVENT): void

  /**
   * Returns the last event sent to the appender.
   * @returns The last LogEvent object sent, or null if none sent yet.
   * @throws ScriptError If the appender instance is not available (not instantiated).
   */
  getLastLogEvent(): LogEvent | null

  /** Returns a string representation of the appender, typically showing the last message sent.
   * @throws ScriptError If the appender instance is not available (not instantiated).
   */
  toString(): string

}

/**
 * Represents a logging interface for capturing and managing log events at various levels.
 * Provides methods for logging messages, querying log state, managing appenders, and exporting logger state.
 * Implementations should ensure thread safety and efficient log event handling.
 * @interface
 */
interface Logger {
  /**
   * Sends an error log event with the provided message to all appenders.
   * @param msg - The error message to log.
    * @throws ScriptError if 
   *          - The singleton instance is not available (not instantiated)
   *          - The logger is configured to exit on error/warning events.
   */
  error(msg: string): void

  /**
   * Sends a warning log event with the provided message to all appenders.
   * @param msg - The warning message to log.
   * @throws ScriptError if 
   *          - The singleton instance is not available (not instantiated)
   *          - The logger is configured to exit on error/warning events.
   */
  warn(msg: string): void

  /**
   * Sends an informational log event with the provided message to all appenders.
   * @param msg - The informational message to log.
   * @throws ScriptError if the singleton instance is not available (not instantiated).
   */
  info(msg: string): void

  /**
   * Sends a trace log event with the provided message to all appenders.
   * @throws ScriptError if the singleton instance is not available (not instantiated).
   * @param msg - The trace message to log.
   */
  trace(msg: string): void

  /**
   * Retrieves an array of all error and warning log events sent.
   * @returns An array of LogEvent objects representing error and warning events.
   * @throws ScriptError if the singleton instance is not available (not instantiated).
   */
  getCriticalEvents(): LogEvent[]

  /**
   * Gets the total number of error log events sent.
   * @throws ScriptError if the singleton instance is not available (not instantiated).
   * @returns The count of error events.
   */
  getErrCnt(): number

  /**
   * Gets the total number of warning log events sent.
   * @returns The count of warning events.
   * @throws ScriptError if the singleton instance is not available (not instantiated).
   */
  getWarnCnt(): number

  /**
   * Gets the current action setting for error/warning events.
   * @returns The action value (e.g., CONTINUE or EXIT).
   * @throws ScriptError if the singleton instance is not available (not instantiated).
   */
  getAction(): number

  /**
   * Gets the current log level setting.
   * @returns The log level value (e.g., OFF, ERROR, WARN, INFO, TRACE).
   * @throws ScriptError if the singleton instance is not available (not instantiated).
   */
  getLevel(): number

  /**
   * Retrieves the array of appenders currently registered with the logger.
   * @returns An array of Appender instances.
   * @throws ScriptError if the singleton instance is not available (not instantiated).
   */
  getAppenders(): Appender[]

  /**
   * Sets the array of appenders for the logger.
   * @param appenders - The array of Appender instances to set.
   * @throws ScriptError if 
   *         - The singleton instance is not available (not instantiated).
   *         - The resulting array doesn't contain unique implementations of Appender.
   *         - appender is null or undefined or has null or undefined elements
   */
  setAppenders(appenders: Appender[]): void

  /**
   * Adds a new appender to the logger.
   * @param appender - The Appender instance to add.
   * @throws ScriptError if 
   *         - The singleton instance is not available (not instantiated).
   *         - The appender is null or undefined.
   *         - The appender is already registered in the logger.
   * @see Logger.setAppenders() for setting multiple appenders at once and for more details
   *    on the validation of the appenders.
   */
  addAppender(appender: Appender): void

  /**
   * Removes an appender from the logger, if the resulting array of appenders appender is not empty.
   * @param appender - The Appender instance to remove.
   * @throws ScriptError if the singleton instance is not available (not instantiated).
   */
  removeAppender(appender: Appender): void

  /**
   * Checks if any error log events have been sent.
   * @returns True if at least one error event has been sent; otherwise, false.
   * @throws ScriptError if the singleton instance is not available (not instantiated).
   */
  hasErrors(): boolean

  /**
   * Checks if any warning log events have been sent.
   * @returns True if at least one warning event has been sent; otherwise, false.
   * @throws ScriptError if the singleton instance is not available (not instantiated).
   */
  hasWarnings(): boolean

  /**
   * Checks if any error or warning log events have been sent.
   * @returns True if at least one error or warning event has been sent; otherwise, false.
   * @throws ScriptError if the singleton instance is not available (not instantiated).
   */
  hasMessages(): boolean

  /**
   * Clears the logger's history of error and warning events and resets counters.
   * @throws ScriptError if the singleton instance is not available (not instantiated).
   */
  clear(): void

  /**
   * Exports the current state of the logger, including level, action, error/warning counts, and critical events.
   * @returns An object containing the logger's state.
   * @throws ScriptError if the singleton instance is not available (not instantiated).
   */
  exportState(): {
    level: string
    action: string
    errorCount: number
    warningCount: number
    criticalEvents: LogEvent[]
  }

  /**
   * Returns a string representation of the logger's state, including level, action, and message counts.
   * @returns A string describing the logger's current state.
   * @throws ScriptError if the singleton instance is not available (not instantiated).
   */
  toString(): string
}

// CLASSES
// --------------------

/**
 * Utility class providing static helper methods for logging operations.
 */
class Utility {
/**Helpder to format the local date as a string. Ouptut in standard format: YYYY-MM-DD HH:mm:ss,SSS
*/
  public static date2Str(date:Date): string {
    const pad = (n: number, width = 2) => n.toString().padStart(width, '0');
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)
      }-${pad(date.getDate())
      } ${pad(date.getHours())
      }:${pad(date.getMinutes())
      }:${pad(date.getSeconds())
      },${pad(date.getMilliseconds(), 3)
      }`
  }

  /** Helper method to check for an empty array. */
  public static isEmptyArray<T>(arr: T[]): boolean {
    return (!Array.isArray(arr) || !arr.length) ? true : false;
  }

  /**
   * Validates a factory function with one argument was well defined. used to validate Log event factory functions.
   * Ensures it is a function with the correct signature for creating log events.
   * @param factory The factory function to validate.
   * @param funName Used to identify the function name in the error message.
   * @param context 
   * @throws ScriptError if the log event factory is not a function or does not have the expected arity (2 parameters).
   */
  static validateFun2Arg(factory: LogEventFactory, funName: string = "LogEventFactory",
    context?: string): void {
    const PREFIX = context ? `[${context}]: ` : ''
    if (typeof factory !== "function") {
      throw new ScriptError(`${PREFIX}Invalid ${funName}: Not a function`)
    }
    if (factory.length !== 2) {
      throw new ScriptError(`${PREFIX}Invalid ${funName}: Must take exactly two arguments (message: string, eventType: LOG_EVENT)`);
    }
  }

}


/**
 * Implements the LogEvent interface, providing a concrete representation of a log event.
 * It includes properties for the event type, message, and timestamp, along with methods to manage
 * the layout used for formatting log events before sending them to appenders.
 */
class LogEventImpl implements LogEvent {
  private readonly _type: LOG_EVENT
  private readonly _message: string
  private readonly _timestamp: Date

  /**
   * Constructs a new LogEventImpl instance.
   * Validates the input parameters to ensure they conform to expected types and constraints.
   * @param type - The type of the log event (from LOG_EVENT enum).
   * @param message - The message to log.
   * @param timestamp - (Optional) The timestamp of the event, defaults to current time.
   * @throws ScriptError if validation fails.
   */
  constructor(message: string, type: LOG_EVENT, timestamp: Date = new Date()) {
    LogEventImpl.validateLogEventAttrs({ type: type, message, timestamp }, "LogEventImpl.constructor")
    this._type = type
    this._message = message
    this._timestamp = timestamp
  }

  /**
   * @returns The event type from LOG_EVENT enum (immutable).
   */
  public get type(): LOG_EVENT { return this._type }

  /**
   * @returns The message of the log event (immutable).
   */
  public get message(): string { return this._message }

  /**
   * @returns The timestamp of the log event (immutable).
   */
  public get timestamp(): Date { return this._timestamp }

  /**
   * Validates if the input object conforms to the LogEvent interface (for any implementation).
   * @throws ScriptError if event is invalid.
   */
  public static validateLogEvent(event: unknown, context?:string): void {
    if (typeof event !== 'object' || event == null) {
      const PREFIX = context ? `[${context}]: ` : ''
      throw new ScriptError(`${PREFIX}LogEvent must be a non-null object.`)
    }
    const e = event as { type?: unknown, message?: unknown, timestamp?: unknown }
    LogEventImpl.validateLogEventAttrs({
      type: e.type,
      message: e.message,
      timestamp: e.timestamp
    })
  }

  /**
   * @returns A string representation of the log event in the format: [timestamp] [type] message.
   *          The timestamp is formatted as YYYY-MM-DD HH:mm:ss,SSS.
   */
  public toString(): string {
    const sDATE = Utility.date2Str(this._timestamp) // Local date as string
    const sType = LogEventImpl.eventTypeToLabel(this.type) // Get the string representation of the type  
    return `[${sDATE}] [${sType}] ${this._message}`
  }

  /**
   * Returns a standardized label for the given log event.
   * @param eventType - The event type from 'LOG_EVENT' enum.
   * @returns A string label, e.g., '[INFO]', '[ERROR]'.
   */
  public static eventTypeToLabel(eventType: LOG_EVENT): string {
    return `${LOG_EVENT[eventType]}`
  }

  /**
   * Validates the raw attributes for a log event.
   * @throws ScriptError if any of the attributes are not valid.
   */
  private static validateLogEventAttrs(attrs: { type: unknown, message: unknown, timestamp: unknown },
    context?:string): void {
    const PREFIX = context ? `[${context}]: ` : ''
    if (typeof attrs.type !== 'number') {
      throw new ScriptError(`${PREFIX}LogEvent.eType='${attrs.type}' property must be a number (LOG_EVENT enum value)`);
    }
    if (!Object.values(LOG_EVENT).includes(attrs.type as LOG_EVENT)) {
      throw new ScriptError(`${PREFIX}LogEvent.type='${attrs.type}' property is not defined in the LOG_EVENT enum.`);
    }
    if (typeof attrs.message !== 'string') {
      throw new ScriptError(`${PREFIX}LogEvent.message='${attrs.message}' property must be a string`);
    }
    if (!(attrs.timestamp instanceof Date)) {
      throw new ScriptError(`${PREFIX}LogEvent.timestamp='${attrs.timestamp}' property must be a Date`);
    }
  }
}


/**
 * Default implementation of the 'Layout' interface.
 * Formats log events into a string using a provided or default formatting function.
 *
 * @remarks
 * - Uses the Strategy Pattern for extensibility: you can pass a custom formatting function (strategy) to the constructor.
 * - All events are validated to conform to the LogEvent interface before formatting.
 * - Throws ScriptError if the event does not conform to the expected LogEvent interface.
 */
class LayoutImpl implements Layout {
  /**
   * Convinience public constant to help users to define a short format for log events. 
   * Formats a log event as a short string as follows '[type] message'. 
   * Defined as a named function to ensure toString() returns the function name.
   */
  public static readonly shortFormatterFun = function shortLayoutFormatterFun(event: LogEvent): string {
    const sType = LogEventImpl.eventTypeToLabel(event.type) // String representation of the type
    return `[${sType}] ${event.message}`
  }

  /**
   * Default formatter function. Created as a named function. Formats a log event as [timestamp] [type] message
   * and timestamp is formatted as YYYY-MM-DD HH:mm:ss,SSS.
   * to ensure toString() returns the function name.
   */
  private static readonly defaultFormatterFun = function defaultLayoutFormatterFun(event: LogEvent): string {
    const sDATE = Utility.date2Str(event.timestamp)         // Local date as string
    const sType = LogEventImpl.eventTypeToLabel(event.type) // String representation of the type
    return `[${sDATE}] [${sType}] ${event.message}`
  }

  /**
   * Function used to convert a LogEvent into a string.
   * Set at construction time; defaults to a simple "[type] message" format.
   */
  private readonly _formatter: (logEvent: LogEvent) => string

  /**
   * Constructs a new LayoutImpl.
   * 
   * @param formatter - Optional. A function that formats a LogEvent as a string.
   *                   If not provided, a default formatter is used: "[timestamp] [type] message".
   * 
   * @remarks Strategy Pattern to allow flexible formatting via formatter function.
   * Pass a custom function to change the formatting logic at runtime.
   * @throws ScriptError if:
   *         - The formatter is not a function or does not have the expected arity (1 parameter).
   *         - The formatter is null or undefined.
   *         - The instance object is undefined or null (if subclassed or mutated in ways that break the interface).
   * @example
   * // Using the default formatter:
   * const layout = new LayoutImpl()
   * // Using a custom formatter for JSON output:
   * const jsonLayout = new LayoutImpl(event => JSON.stringify(event))
   * // Using a formatter for XML output:
   * const xmlLayout = new LayoutImpl(event =>
   *   `<log><type>${event.type}</type><message>${event.message}</message></log>`
   * )
   * // Using a shorter format [type] [message]:
   * const shortLayout = new LayoutImpl(e => `[${LOG_EVENT[e.type]}] ${e.message}`)
   * // Using a custom formatter with a named function, so in toString() shows the name of the formatter.
   * let shortLayoutFun: Layout = new LayoutImpl(
   *   function shortLayoutFun(e:LogEvent):string{return `[${LOG_EVENT[e.type]}] ${e.message}`})
   */
  constructor(formatter?: (event: LogEvent) => string) {
    this._formatter = formatter ?? LayoutImpl.defaultFormatterFun
    LayoutImpl.validateLayout(this as unknown, "LayoutImpl.constructor")
  }

  /**
   * Returns the current formatter function.
   * @returns The formatter function.
   */
  public getFormatter(): (logEvent: LogEvent) => string {
    return this._formatter
  }

  /**
   * Formats the given log event as a string.
   * @param event - The event to format.
   * @returns A string representation of the log event.
   * @throws ScriptError if the event does not conform to the LogEvent interface.
   */
  public format(event: LogEvent): string {
    LogEventImpl.validateLogEvent(event)
    return this._formatter(event)
  }

  /**
   * @returns A string representation of the layout.
   *          If the formatter is a function, it returns the name of the function.
   */
  public toString(): string {
    const formatterName = this._formatter?.name || "anonymous"
    return `LayoutImpl { formatter: [Function: ${formatterName}] }`
  }
/**
   * Asserts that the provided object implements the Layout interface.
   * @param layout - The object to validate.
   * @throws ScriptError if
   *          - layout is null or undefined
   *          - the format method is not a function or doesn't have the expected arity (one parameter).
   */
  public static validateLayout(layout: unknown, context?: string): asserts layout is Layout {
    const PREFIX = context ? `[${context}]:` : ''
    if (layout == null) {
      throw new ScriptError(`${PREFIX}Invalid Layout: layout object is null or undefined`)
    }
    const maybeLayout = layout as Record<string, unknown>
    if (typeof maybeLayout.format !== "function") {
      throw new ScriptError(`${PREFIX}Invalid Layout: the 'format' method is missing or not a function`)
    }
    if (maybeLayout.format.length !== 1) {
      throw new ScriptError(`${PREFIX}Invalid Layout: 'format' function should take exactly one argument (logEvent: LogEvent)`)
    }
  }

}

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
   * Utility method to rethrow the deepest original cause if present,
   * otherwise rethrows this 'ScriptError' itself.
   * Useful for deferring a controlled exception and then
   * surfacing the root cause explicitly.
   */
  public rethrowCauseIfNeeded(): never {
    if (this.cause instanceof ScriptError && typeof this.cause.rethrowCauseIfNeeded === "function") {
      // Recursively rethrow the root cause if nested ScriptError
      this.cause.rethrowCauseIfNeeded()
    } else if (this.cause) {
      // Rethrow the immediate cause if not a ScriptError
      throw this.cause
    }
    // No cause, throw self
    throw this
  }

  /** Override toString() method.
   * @returns The name and the message on the first line, then 
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
 * 
 * It relies on a LogEventFactory function to create log events, enabling flexible
 * and customizable event creation strategies. The LogEventFactory is validated on
 * construction to ensure it meets the expected signature. This design allows users
 * to supply custom event creation logic if needed.
 * 
 * Appenders such as 'ConsoleAppender' and 'ExcelAppender' should extend this class
 * to inherit consistent logging behavior and event creation via the LogEventFactory.
 */
abstract class AbstractAppender implements Appender {
  /** Formats the log message based on the event type. Singleton pattern. */
  private static _layout: Layout | null = null // Static layout shared by all events, lazy initialized
  private _lastLogEvent: LogEvent | null = null // The last event sent by the appender
  private _logEventFactory: LogEventFactory
  private static readonly _logEventFactoryFun: LogEventFactory 
    = function logEventFactoryFun (message: string, eventType: LOG_EVENT) {
      return new LogEventImpl(message, eventType)
    } // Default factory function to create LogEvent instances

/**
 * Constructs a new AbstractAppender instance.
 * @param logEventFactory - Optional. A factory function to create LogEvent instances.
 *                          Must have the signature (message: string, eventType: LOG_EVENT) => LogEvent.
 *                          If not provided, a default factory function is used.
 * @throws ScriptError if the log event factory is not a valid function with the expected signature.
 * @example
 * // Example: Custom LogEvent with a custom date format field
 * class MyCustomLogEvent extends LogEventImpl {
 *   public readonly dateFormat: string;
 *   constructor(message: string, eventType: LOG_EVENT, timestamp: number, dateFormat: string) {
 *     super(message, eventType, timestamp)
 *     this.dateFormat = dateFormat
 *   }
 * }
 * // Custom factory that injects a specific date format into each event
 * const customFactory = (msg: string, eventType: LOG_EVENT) => {
 *   const dateFormat = "YYYY-MM-DD HH:mm:ss"; // your custom format
 *   return new MyCustomLogEvent(msg, eventType, Date.now(), dateFormat);
 * }
 * const appender = new MyAppender(customFactory);
 */
  protected constructor(logEventFactory: LogEventFactory = AbstractAppender._logEventFactoryFun) {
    this._logEventFactory = logEventFactory
    // Validate the factory function
    Utility.validateFun2Arg(this._logEventFactory, "logEventFactory", "AbstractAppender.constructor") 
  }

  /**
   * @returns The layout associated to all events. Used to format the log event before sending it to the appenders. 
   * If the layout was not set, it returns a default layout (lazy initialization). The layout is shared by all events
   * and all appenders, so it is static.
   * @remarks Static, shared by all log events. Singleton.
   */
  public static getLayout(): Layout {
    if (!AbstractAppender._layout) {
      AbstractAppender._layout = new LayoutImpl() // Default layout if not set
    }
    return AbstractAppender._layout
  }

  /**
   * Sets the layout associated to all events, the layout is assigned only if it was not set before.
   * @param layout - The layout to set.
   */
  public static setLayout(layout: Layout): void {
    LayoutImpl.validateLayout(layout)
    if (!AbstractAppender._layout) {
      AbstractAppender._layout = layout
    }
  }

  /**
   * Log a message or event.
   * @param arg1 - LogEvent or message string.
   * @param arg2 - LOG_EVENT, only required if arg1 is a string.
   */
  public log(arg1: LogEvent | string, arg2?: LOG_EVENT): void {
    if (typeof arg1 === "string") {
      if (arg2 === undefined) {
        throw new ScriptError("event type must be provided when logging with a message string.")
      }
      // Uses the LogEventFactory type for construction.
      const event =  this._logEventFactory(arg1, arg2)
      this.sendEvent(event)
    } else {
      this.sendEvent(arg1)
    }
  }

  /**
   * @returns The last log event sent by the appender.
   */
  public getLastLogEvent(): LogEvent | null { 
    return this._lastLogEvent
  }

  // #TEST-ONLY-START
  /** Sets to null the static layout, useful for running different scenarios.*/
  public static clearLayout(): void {
    AbstractAppender._layout = null // Force full re-init
  }
  // #TEST-ONLY-END

  /**
   * Returns a string representation of the appender.
   * It includes the appender name, the log event factory name, and the layout name.
   * @returns A string representation of the appender.
   * @throws ScriptError if 
   *         - The appender instance is not available (not instantiated).
   *         - The layout is not set or is invalid.
   *         - The log event factory is not a function or does not have the expected arity (2 parameters).
   */
  public toString(): string {
    const name = this.constructor.name
    const logEventFactoryName = this._logEventFactory.name || "anonymous"
    LayoutImpl.validateLayout(AbstractAppender.getLayout()) // Ensure layout is valid
    const layoutName = AbstractAppender.getLayout().toString() // Get the layout name
    return `${name}: {LogEventFactory: '${logEventFactoryName}', Layout: '${layoutName}'}`
  }

  /**
   * Send the event to the appropriate destination. The event is storeed based on the 
   * @param event - The log event to be sent.
   * @throws ScriptError if 
   *         - The event is not a valid LogEvent.
   */
  protected abstract sendEvent(event: LogEvent): void

  /** Stores the last log event sent to the appender 
   * @param event - The log event to store.
   * @throws ScriptError if the event is not a valid LogEvent
  */
  protected setLastLogEvent(event: LogEvent): void {
    LogEventImpl.validateLogEvent(event) // Validate the event
    this._lastLogEvent = event // Set the last log event
  }

}

/**
 * Singleton appender that logs messages to the Office Script console. It is used as default appender, 
 * if no other appender is defined. The content of the message event sent can be customized via
 * any Layout implementation, but by default it uses the LayoutImpl.
 * Usage:
 * - Call ConsoleAppender.getInstance() to get the appender
 * - Automatically used if no other appender is defined
 * @example:
 * // Add console appender to the Logger
 * Logger.addAppender(ConsoleAppender.getInstance())
 */
class ConsoleAppender extends AbstractAppender implements Appender  {
  private static _instance: ConsoleAppender | null = null // Instance of the singleton pattern

  /**
   * Private constructor to prevent user instantiation.
   * Initializes the appender without a layout, using the default layout if not set.
   */
  private constructor() {
    super() // Call the parent constructor without layout
  }

  /**Override toString method. Show the last message event sent
   * @throws ScriptError If the singleton was not instantiated
   */
  public toString(): string {
    ConsoleAppender.validateInstance()
    const name = this.constructor.name
    const lastEvent = this.getLastLogEvent() ?? null
    const formatted = lastEvent ? AbstractAppender.getLayout().format(lastEvent):""
    return `${name}: {Last log event: '${formatted}'}`
  }

  /**
   * @returns The singleton instance of the class. 
   * If it was not instantiated, it creates a new instance with the provided layout.
   * It sets a default layout if it was not set before.
   */
  public static getInstance(): ConsoleAppender {
    if(!ConsoleAppender._instance) {
      ConsoleAppender._instance = new ConsoleAppender()
      if (!AbstractAppender.getLayout()) { // Set the default layout if not set
        AbstractAppender.setLayout(new LayoutImpl()) // Default layout if not set
      }
    }
    return ConsoleAppender._instance
  }

  // #TEST-ONLY-START
  /** Sets to null the singleton instance, useful for running different scenarios.
   * Warning: Mainly intended for testing purposes. The state of the singleton will be lost.
   * @example There is no way to empty the last message sent after the instance was created unless
   * you reset it.
   * appender:ConsoleAppender = ConsoleAppender.getInstance()
   * appender = ConsoleAppender.log("info event", LOG_EVENT.INFO)
   * appender.getLastLogEvent().message     // Output: info event"
   * appender.clearInstance()               // clear the singleton
   * appender = getInstance()               // restart the singleton
   * appender.getLastLogEvent().message     // Output: ""
   */
  public static clearInstance(): void {
    ConsoleAppender._instance = null // Force full re-init
  }
  // #TEST-ONLY-END

  /** @internal
   * Common safeguard method, where calling initIfNeeded doesn't make sense.
   * @throws ScriptError In case the singleton was not initialized.
   */
  private static validateInstance():void {
    if (!ConsoleAppender._instance) {
      const MSG = `In '${ConsoleAppender.name}' class a singleton instance can't be undefined or null. Please invoke getInstance first`
      throw new ScriptError(MSG)
    }
  }

  /**
   * Internal method to send the event to the console.
   * @param event - The log event to output.
   */
  protected sendEvent(event: LogEvent): void {
    LogEventImpl.validateLogEvent(event)
    ConsoleAppender.validateInstance()
    this.setLastLogEvent(event)
    const MSG = AbstractAppender.getLayout().format(event)
    console.log(MSG)
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
  private static readonly _DEFAULT_COLOR_MAP = Object.freeze({
    [LOG_EVENT.ERROR]: "9c0006",  // RED
    [LOG_EVENT.WARN]: "ed7d31",   // ORANGE
    [LOG_EVENT.INFO]: "548235",   // GREEN
    [LOG_EVENT.TRACE]: "7f7f7f"   // GRAY
  } as const);

  // Static map to associate LOG_EVENT types with font input argument from getInstance()
  private static readonly FONT_LABEL_MAP = Object.freeze({
    [LOG_EVENT.ERROR]: "errFont",
    [LOG_EVENT.WARN]: "warnFont",
    [LOG_EVENT.INFO]: "infoFont",
    [LOG_EVENT.TRACE]: "traceFont",
  } as const);

  /**
   * Instance-level color map for current appender configuration.
   * Maps LOG_EVENT types to hex color strings.
   */
  private readonly colorMap: Record<LOG_EVENT, string>;

  /* Regular expression to validate hexadecimal colors */
  private static readonly HEX_REGEX = Object.freeze(/^#?[0-9A-Fa-f]{6}$/);

  private static _instance: ExcelAppender | null = null; // Instance of the singleton pattern
  private readonly _msgCellRng: ExcelScript.Range;

  // Private constructor to prevent user invocation
  private constructor(
    msgCellRng: ExcelScript.Range,
    errFont: string = ExcelAppender._DEFAULT_COLOR_MAP[LOG_EVENT.ERROR],
    warnFont: string = ExcelAppender._DEFAULT_COLOR_MAP[LOG_EVENT.WARN],
    infoFont: string = ExcelAppender._DEFAULT_COLOR_MAP[LOG_EVENT.INFO],
    traceFont: string = ExcelAppender._DEFAULT_COLOR_MAP[LOG_EVENT.TRACE]
  ) {
    super();
    this._msgCellRng = msgCellRng;
    this._msgCellRng.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    this.clearCellIfNotEmpty(); // Clear the cell if it has a value
    this.colorMap = {
      [LOG_EVENT.ERROR]: errFont,
      [LOG_EVENT.WARN]: warnFont,
      [LOG_EVENT.INFO]: infoFont,
      [LOG_EVENT.TRACE]: traceFont
    };
    // Set the default layout if not set
    if (!AbstractAppender.getLayout()) {
      AbstractAppender.setLayout(new LayoutImpl()); // Default layout if not set
    }
  }

  /**
   * Returns the singleton ExcelAppender instance, creating it if it doesn't exist.
   * On first call, requires a valid single cell Excel range to display log messages and optional
   * color customizations for different log events (LOG_EVENT). Subsequent calls ignore parameters
   * and return the existing instance.
   * @param msgCellRng - Excel range where log messages will be written. Must be a single cell and
   * not null or undefined.
   * @param errFont - Hex color code for error messages (default: "9c0006" red).
   * @param warnFont - Hex color code for warnings (default: "ed7d31" orange).
   * @param infoFont - Hex color code for info messages (default: "548235" green).
   * @param traceFont - Hex color code for trace messages (default: "7f7f7f" gray).
   * @returns The singleton ExcelAppender instance.
   * @throws ScriptError if msgCellRng was not defined or if the range covers multiple cells
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
    errFont: string = ExcelAppender._DEFAULT_COLOR_MAP[LOG_EVENT.ERROR],
    warnFont: string = ExcelAppender._DEFAULT_COLOR_MAP[LOG_EVENT.WARN],
    infoFont: string = ExcelAppender._DEFAULT_COLOR_MAP[LOG_EVENT.INFO],
    traceFont: string = ExcelAppender._DEFAULT_COLOR_MAP[LOG_EVENT.TRACE]
  ): ExcelAppender {
    if (!ExcelAppender._instance) {
      if (!msgCellRng || !msgCellRng.setValue) {
        const MSG = `${ExcelAppender.name} requires a valid ExcelScript.Range for input argument msgCellRng.`;
        throw new ScriptError(MSG);
      }
      if (msgCellRng.getCellCount() != 1) {
        const MSG = `${ExcelAppender.name} requires input argument msgCellRng represents a single Excel cell.`;
        throw new ScriptError(MSG);
      }
      // Checking valid hexadecimal color
      ExcelAppender.assertColor(errFont, "error");
      ExcelAppender.assertColor(warnFont, "warning");
      ExcelAppender.assertColor(infoFont, "info");
      ExcelAppender.assertColor(traceFont, "trace");
      ExcelAppender._instance = new ExcelAppender(msgCellRng, errFont, warnFont, infoFont, traceFont);
    }
    return ExcelAppender._instance;
  }

  // #TEST-ONLY-START
  /**
   * Sets to null the singleton instance, useful for running different scenarios.
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
    ExcelAppender._instance = null;
  }
  // #TEST-ONLY-END

  /**
   * Shows instance configuration plus last message sent by the appender
   * @throws ScriptError, if the singleton was not instantiated.
   */
  public toString(): string {
    ExcelAppender.validateInstance();
    const NAME = this.constructor.name;
    const MSG_CELL_RNG = this._msgCellRng.getAddress();
    const lastEventObj = AbstractAppender.getLayout().format(this.getLastLogEvent());
    const LAST_EVENT = lastEventObj ?? "(no event sent)";
    // Present the color map in the output as "event colors"
    const EVENT_COLORS = Object.entries(this.colorMap).map(
      ([key, value]) => [
        ExcelAppender.FONT_LABEL_MAP[Number(key) as LOG_EVENT],
        value
      ]
    );
    return `${NAME}: {Message Range: "${MSG_CELL_RNG}", Event fonts: {${EVENT_COLORS}}, Last log event: "${LAST_EVENT}"}`;
  }

  /**
   * Sets the value of the cell, with the event message, using the font defined for the event type,
   * if not font was defined it doesn't change the font of the cell.
   * @param event a value from enum LOG_EVENT.
   * @throws ScriptError in case event is not a valid LOG_EVENT enum value.
   */
  protected sendEvent(event: LogEvent): void {
    ExcelAppender.validateInstance();
    LogEventImpl.validateLogEvent(event);
    this.clearCellIfNotEmpty(); // Clear the cell if it has a value
    const FONT = this.colorMap[event.type] ?? null;
    if (FONT) {
      this._msgCellRng.getFormat().getFont().setColor(FONT);
    }
    const MSG = AbstractAppender.getLayout().format(event);
    this._msgCellRng.setValue(MSG);
    this._msgCellRng.getValue(); // Explicitly access the cell to ensure it commits the update
    this.setLastLogEvent(event);
  }

  // Common safeguard method
  private static validateInstance() {
    if (!ExcelAppender._instance) {
      const MSG = `In '${ExcelAppender.name}' class a singleton instance can't be undefined or null. Please invoke getInstance first`;
      throw new ScriptError(MSG);
    }
  }

  // Validate color is a valid hexadecimal color
  private static assertColor(color: string, name: string): void {
    const match = ExcelAppender.HEX_REGEX.test(color);
    if (!match) {
      const MSG = `The input value '${color}' color for '${name}' event is not a valid hexadecimal color. Please enter a value that matches the following regular expression: '${ExcelAppender.HEX_REGEX.toString()}'`;
      throw new ScriptError(MSG);
    }
  }

  /**
   * Clears the message cell only if it is not empty.
   */
  private clearCellIfNotEmpty(): void {
    const value = this._msgCellRng.getValue();
    if (value !== undefined && value !== null && value !== "") {
      this._msgCellRng.clear(ExcelScript.ClearApplyTo.contents);
    }
  }
}

/**
 * Singleton class that manages application logging through appenders.
 * Supports the following log events: ERROR, WARN, INFO, TRACE (LOG_EVENT enum).
 * Supports the level of information (verbose) to show via Logger.LEVEL: OFF, ERROR, WARN, INFO, TRACE.
 * If the level of information (LEVEL) is OFF, no log events will be sent to the appenders.
 * Supports the action to take in case of ERROR, WARN log events: the script can
 * continue ('Logger.ACTION.CONTINUE'), or abort ('Logger.ACTION.EXIT'). Such actions only take effect
 * if the LEVEL is not Logger.LEVEL.OFF.
 * Allows defining appenders, controlling the channels the events are sent to.
 * Collects error/warning sent by the appenders via getMessages().
 * 
 * Usage:
 * - Initialize with Logger.getInstance(level, action)
 * - Add one or more appenders (e.g. ConsoleAppender, ExcelAppender)
 * - Use Logger.error(), warn(), info(), or trace() to log
 * 
 * Features:
 * - If no appender is added, ConsoleAppender is used by default
 * - Logs are routed through all registered appenders
 * - Collects a summary of error/warning messages and counts
 */
class LoggerImpl implements Logger {
  // Constants
  public static ACTION = Object.freeze({
    CONTINUE: 0, // In case of error/warning log events, the script continues
    EXIT: 1,     // In case of error/warning log event, throws a ScriptError error
  } as const)

  /* Generates the same sequence as LOG_EVENT, but adding the zero case with OFF. It ensures the numeric values
  match the values of LOG_EVENT. Note: enum can't be defined inside a class */
  public static readonly LEVEL = Object.freeze(Object.assign({ OFF: 0 }, LOG_EVENT))

  // Equivalent labels from LEVEL
  private static readonly LEVEL_LABELS = Object.entries(LoggerImpl.LEVEL).reduce((acc, [key, value]) => {
    acc[value] = key;
    return acc;
  }, {} as Record<string, string>)

  // Equivalent labels from ACTION
  private static readonly ACTION_LABELS = Object.entries(LoggerImpl.ACTION).reduce((acc, [key, value]) => {
    acc[value] = key;
    return acc;
  }, {} as Record<number, string>)

  /**Default factory for LogEvent instances */
  private static readonly DEFAULT_EVENT_FACTORY: LogEventFactory =
    (msg: string, eventType: LOG_EVENT) => new LogEventImpl(msg, eventType)

  // Attributes
  private static _instance: LoggerImpl | null = null; // Instance of the singleton pattern
  private static readonly DEFAULT_LEVEL = LoggerImpl.LEVEL.WARN
  private static readonly DEFAULT_ACTION = LoggerImpl.ACTION.EXIT

  private readonly _level: typeof LoggerImpl.LEVEL[keyof typeof LoggerImpl.LEVEL] = LoggerImpl.DEFAULT_LEVEL;
  private readonly _action: typeof LoggerImpl.ACTION[keyof typeof LoggerImpl.ACTION] = LoggerImpl.DEFAULT_ACTION;
  private _criticalEvents: LogEvent[] = []; // Collects all ERROR and WARN events only
  private _errCnt = 0;   // Counts the number of error events found
  private _warnCnt = 0;  // Counts the number of warning events found
  private _appenders: Appender[] = []; // List of appenders
  private readonly _logEventFactory: LogEventFactory // Factory function to create LogEvent instances 

  private constructor(
    level: typeof LoggerImpl.LEVEL[keyof typeof LoggerImpl.LEVEL] = LoggerImpl.DEFAULT_LEVEL,
    action: typeof LoggerImpl.ACTION[keyof typeof LoggerImpl.ACTION] = LoggerImpl.DEFAULT_ACTION,
    logEventFactory: LogEventFactory = LoggerImpl.DEFAULT_EVENT_FACTORY
  ) {
    this._action = action
    this._level = level
    Utility.validateFun2Arg(logEventFactory, "logEventFactory", "LoggerImpl.constructor")
    this._logEventFactory = logEventFactory
  }

  // Getters
  /** 
   * @returns An array with error and warning event messages only.
   * @throws ScriptError If the singleton was not instantiated.
   */
  public getCriticalEvents(): LogEvent[] {
    LoggerImpl.validateInstance()
    return LoggerImpl._instance._criticalEvents
  }

  /** 
   * @returns Total number of error message events sent to the appenders.
   * @throws ScriptError If the singleton was not instantiated.
   */
  public getErrCnt(): number {
    LoggerImpl.validateInstance()
    return LoggerImpl._instance._errCnt
  }

  /** 
   * @returns Total number of warning events sent to the appenders.
   * @throws ScriptError If the singleton was not instantiated.
   */
  public getWarnCnt(): number {
    LoggerImpl.validateInstance()
    return LoggerImpl._instance._warnCnt
  }

  /** 
   * @returns The action to take in case of errors or warning log events.
   * @throws ScriptError If the singleton was not instantiated.
   */
  public getAction(): typeof LoggerImpl.ACTION[keyof typeof LoggerImpl.ACTION] {
    LoggerImpl.validateInstance()
    return LoggerImpl._instance._action
  }

  /** 
   * Returns the level of verbosity allowed in the Logger. The levels are incremental, i.e.
   * it includes all previous levels. For example: Logger.WARN includes warnings and errors since
   * Logger.ERROR is lower.
   * @returns The current log level.
   * @throws ScriptError If the singleton was not instantiated.
   */
  public getLevel(): typeof LoggerImpl.LEVEL[keyof typeof LoggerImpl.LEVEL] {
    LoggerImpl.validateInstance()
    return LoggerImpl._instance._level
  }

  /**
   * @returns Array with appenders subscribed to the Logger.
   * @throws ScriptError If the singleton was not instantiated.
   */
  public getAppenders(): Appender[] {
    LoggerImpl.validateInstance();
    return LoggerImpl._instance._appenders
  }

  // Setters
  /** 
   * Sets the array of appenders with the input argument appenders.
   * @param appenders Array with all appenders to set.
   * @throws ScriptError If the singleton was not instantiated,
   *                     if appenders is null or undefined, or contains
   *                     null or undefined entries,
   *                     or if the appenders to add are not unique
   *                     by appender class. See JSDoc from addAppender.
   * @see addAppender
   */
  public setAppenders(appenders: Appender[]) {
    LoggerImpl.validateInstance()
    LoggerImpl.assertUniqueAppenderTypes(appenders)
    LoggerImpl._instance._appenders = appenders
  }

  /**
   * Adds an appender to the list of appenders.
   * @param appender The appender to add.
   * @throws ScriptError If the singleton was not instantiated,
   *                     if the input argument is null or undefined,
   *                     or if it breaks the class uniqueness of the appenders.
   *                     All appenders must be from a different implementation of the Appender class.
   * @see setAppenders
   */
  public addAppender(appender: Appender): void {
    LoggerImpl.validateInstance();
    if (!appender) { // It must be a valid appender
      const MSG = `You can't add an appender that is null or undefined in the '${LoggerImpl.name}' class.`;
      throw new ScriptError(MSG);
    }
    const newAppenders = [...LoggerImpl._instance._appenders, appender]
    LoggerImpl.assertUniqueAppenderTypes(newAppenders)
    LoggerImpl._instance._appenders.push(appender)
  }

  /**
   * Returns the singleton Logger instance, creating it if it doesn't exist.
   * If the Logger is created during this call, the provided 'level' and 'action'
   * parameters initialize the log level and error-handling behavior.
   * Subsequent calls ignore these parameters and return the existing instance.
   * @param level Initial log level (default: Logger.LEVEL.WARN). Controls verbosity.
   *                Sends events to the appenders up to the defined level of verbosity.
   *                The level of verbosity is incremental, except for value
   *                Logger.LEVEL.OFF, which suppresses all messages sent to the appenders.
   *                For example: Logger.LEVEL.INFO allows sending errors, warnings, and information events,
   *                but excludes trace events.
   * @param action Action on error/warning (default: Logger.ACTION.EXIT).
   *                 Determines if the script should continue or abort.
   *                 If the value is Logger.ACTION.EXIT, throws a ScriptError exception,
   *                 i.e. aborts the Script. If the action is Logger.ACTION.CONTINUE, the
   *                 script continues.
   * @param logEventFactory Optional. Factory function for creating LogEvent instances.
   * @returns The singleton Logger instance.
   * @throws ScriptError If the level input value was not defined in Logger.LEVEL.
   * @example
   * ```ts
   * // Initialize logger at INFO level, continue on errors/warnings
   * const logger = Logger.getInstance(Logger.LEVEL.INFO, Logger.ACTION.CONTINUE);
   * // Subsequent calls ignore parameters, return the same instance
   * const sameLogger = Logger.getInstance(Logger.LEVEL.ERROR, Logger.ACTION.EXIT);
   * Logger.info("Starting the Script"); // send this message to all appenders
   * Logger.trace("Step one"); // Doesn't send because of Logger.LEVEL value: INFO
   * ```
   * @see AbstractAppender.constructor for more information on how to use the logEventFactory. 
   */
  public static getInstance(
    level: typeof LoggerImpl.LEVEL[keyof typeof LoggerImpl.LEVEL] = LoggerImpl.DEFAULT_LEVEL,
    action: typeof LoggerImpl.ACTION[keyof typeof LoggerImpl.ACTION] = LoggerImpl.DEFAULT_ACTION,
    logEventFactory: LogEventFactory = LoggerImpl.DEFAULT_EVENT_FACTORY
  ): LoggerImpl {
    if (!LoggerImpl._instance) {
      LoggerImpl.assertValidLevel(level);
      Utility.validateFun2Arg(logEventFactory, "logEventFactory", "LoggerImpl.getInstance");
      LoggerImpl._instance = new LoggerImpl(level, action, logEventFactory);
    }
    return LoggerImpl._instance;
  }

  // #TEST-ONLY-START
  /** 
   * Sets the singleton instance to null, useful for running different scenarios.
   * Warning: Mainly intended for testing purposes. The state of the singleton will be lost.
   * @example
   * ```ts
   * // Testing how the logger works with default configuration, and then changing the configuration.
   * // Since the class doesn't define setter methods to change the configuration, you can use
   * // clearInstance to reset the singleton and instantiate it with different configuration.
   * // Testing default configuration
   * Logger.getInstance(); // LEVEL: WARN, ACTION: EXIT
   * logger.error("error event"); // Output: "error event" and ScriptError
   * // Now we want to test with the following configuration: Logger.LEVEL:WARN, Logger.ACTION:CONTINUE
   * Logger.clearInstance(); // Clear the singleton
   * Logger.getInstance(LEVEL.WARN, ACTION.CONTINUE);
   * Logger.error("error event"); // Output: "error event" (no ScriptError was thrown)
   * ```
   */
  public static clearInstance(): void {
    LoggerImpl._instance = null; // Force full re-init
  }
  // #TEST-ONLY-END

  /**
 * If the list of appenders is not empty, removes the appender from the list.
 * @param appender The appender to remove.
 * @throws ScriptError If the singleton was not instantiated.
 */
  public removeAppender(appender: Appender): void {
    LoggerImpl.validateInstance();
    const appenders = LoggerImpl._instance._appenders;
    if (!Utility.isEmptyArray(appenders)) {
      const index = LoggerImpl._instance._appenders.indexOf(appender);
      if (index > -1) {
        LoggerImpl._instance._appenders.splice(index, 1); // Remove one element at index
      }
    }
  }

  /**
   * Sends an error log message to all appenders if the level allows it.
   * The level has to be greater than or equal to Logger.LEVEL.ERROR to send this event to the appenders.
   * After the message is sent, it updates the error counter.
   * @remarks
   * If no singleton was defined, it does lazy initialization with default configuration.
   * If no appender was defined, it does lazy initialization to ConsoleAppender.
   * @throws ScriptError Only if level is greater than Logger.LEVEL.OFF and the action is Logger.ACTION.EXIT.
   */
  public error(msg: string): void {
    this.log(msg, LOG_EVENT.ERROR)
  }

  /**
   * Sends a warning event message to the appenders if the level allows it.
   * The level has to be greater than or equal to Logger.LEVEL.WARN to send this event to the appenders.
   * After the message is sent, it updates the warning counter.
   * @remarks
   * If no singleton was defined, it does lazy initialization with default configuration.
   * If no appender was defined, it does lazy initialization to ConsoleAppender.
   * @throws ScriptError Only if level (see getInstance) is greater than Logger.LEVEL.ERROR and the action is Logger.ACTION.EXIT.
   */
  public warn(msg: string): void {
    this.log(msg, LOG_EVENT.WARN)
  }

  /**
   * Sends an info events message to the appenders if the level allows it.
   * The level has to be greater than or equal to Logger.LEVEL.INFO to send this event to the appenders.
   * @remarks
   * If no singleton was defined, it does lazy initialization with default configuration.
   * If no appender was defined, it does lazy initialization to ConsoleAppender.
   */
  public info(msg: string): void {
    this.log(msg, LOG_EVENT.INFO)
  }

  /**
   * Sends a trace events message to the appenders if the level allows it.
   * The level has to be greater than or equal to Logger.LEVEL.TRACE to send this event to the appenders.
   * @remarks
   * If no singleton was defined, it does lazy initialization with default configuration.
   * If no appender was defined, it does lazy initialization to ConsoleAppender.
   */
  public trace(msg: string): void {
    this.log(msg, LOG_EVENT.TRACE)
  }

  /**
   * @returns true if an error log event was sent to the appenders, otherwise false.
   * @throws ScriptError If the singleton was not instantiated.
   */
  public hasErrors(): boolean {
    LoggerImpl.validateInstance();
    return LoggerImpl._instance._errCnt > 0;
  }

  /**
   * @returns true if a warning log event was sent to the appenders, otherwise false.
   * @throws ScriptError If the singleton was not instantiated.
   */
  public hasWarnings(): boolean {
    LoggerImpl.validateInstance();
    return LoggerImpl._instance._warnCnt > 0;
  }

  /**
   * @returns true if some error or warning event has been sent by the appenders, otherwise false.
   * @throws ScriptError If the singleton was not instantiated.
   */
  public hasMessages(): boolean {
    LoggerImpl.validateInstance();
    return LoggerImpl._instance._criticalEvents.length > 0;
  }

  /**
   * Resets the Logger history, i.e., state (errors, warnings, message summary). It doesn't reset the appenders.
   * @throws ScriptError If the singleton was not instantiated.
   */
  public clear(): void {
    LoggerImpl.validateInstance();
    LoggerImpl._instance._criticalEvents = [];
    LoggerImpl._instance._errCnt = 0;
    LoggerImpl._instance._warnCnt = 0;
  }

  /**
   * Serializes the current state of the logger to a plain object, useful for
   * capturing logs and metrics for post-run analysis.
   * For testing/debugging: Compare expected vs actual logger state.
   * For persisting logs into Excel, JSON, or another external system.
   * @throws ScriptError If the singleton was not instantiated.
   * @returns A structure with key information about the logger, such as:
   *        level, action, errorCount, warningCount, criticalEvents.
   */
  public exportState(): {
    level: string,
    action: string,
    errorCount: number,
    warningCount: number,
    criticalEvents: LogEvent[]
  } {
    LoggerImpl.validateInstance();
    const levelKey = Object.keys(LoggerImpl.LEVEL).find(k => LoggerImpl.LEVEL[k as keyof typeof LoggerImpl.LEVEL] === LoggerImpl._instance._level);
    const actionKey = Object.keys(LoggerImpl.ACTION).find(k => LoggerImpl.ACTION[k as keyof typeof LoggerImpl.ACTION] === LoggerImpl._instance._action);

    return {
      level: levelKey ?? "UNKNOWN",
      action: actionKey ?? "UNKNOWN",
      errorCount: LoggerImpl._instance._errCnt,
      warningCount: LoggerImpl._instance._warnCnt,
      criticalEvents: [...LoggerImpl._instance._criticalEvents]
    };
  }

  /**
   * Override toString method.
   * @throws ScriptError If the singleton was not instantiated.
   */
  public toString(): string {
    LoggerImpl.validateInstance();
    const NAME = this.constructor.name;
    const levelTk = Object.keys(LoggerImpl.LEVEL).find(key =>
      LoggerImpl.LEVEL[key as keyof typeof LoggerImpl.LEVEL] === this._level);
    const actionTk = Object.keys(LoggerImpl.ACTION).find(key =>
      LoggerImpl.ACTION[key as keyof typeof LoggerImpl.ACTION] === this._action);
    return `${NAME}: {Level: "${levelTk}", Action: "${actionTk}", Error Count: "${LoggerImpl._instance._errCnt}", Warning Count: "${LoggerImpl._instance._warnCnt}"}`;
  }

  /**
 * Routes the given log event message to all registered appenders.
 * Behavior:
 * - If the singleton was not instantiated, it instantiates it with default configuration (lazy initialization)
 *   to avoid sending unnecessary errors.
 * - If no appenders are defined, a 'ConsoleAppender' is automatically created and added.
 * - The message is only dispatched if the current log level allows it.
 * - If the event is of type 'ERROR' or 'WARN', it is recorded internally and counted.
 * - If the configured action is 'Logger.ACTION.EXIT', a 'ScriptError' is thrown for errors and warnings.
 * @throws ScriptError In case the action defined for the logger is Logger.ACTION.EXIT and the event type
 *      is LOG_EVENT.ERROR or LOG_EVENT.WARN.
 */
private log(msg: string, eventType: LOG_EVENT): void {
  LoggerImpl.initIfNeeded() // lazy initialization of the singleton with default parameters
  const SEND_EVENTS = LoggerImpl._instance._level !== LoggerImpl.LEVEL.OFF
  if (Utility.isEmptyArray(this.getAppenders())) {
    this.addAppender(ConsoleAppender.getInstance()) // lazy initialization at least the basic appender
  }
  if (SEND_EVENTS) { // Sends events through all appenders
    if (LoggerImpl._instance._level >= eventType) { // only if the verbose level allows it
      // Create the event using the factory
      const event = this._logEventFactory(msg, eventType) // same event for all appenders
      for (const appender of LoggerImpl._instance._appenders) { // sends to all appenders
        appender.log(event) // Now use the LogEvent object
      }
    }
  }

  if (SEND_EVENTS && (eventType <= LOG_EVENT.WARN)) { // Only collects errors or warnings event messages
    // Updating the counter
    if (eventType === LOG_EVENT.ERROR) ++LoggerImpl._instance._errCnt
    if (eventType === LOG_EVENT.WARN) ++LoggerImpl._instance._warnCnt
    // Updating the message. Assumes first appender is representative (message for all appenders are the same)
    const appender = LoggerImpl._instance._appenders[0]
    const lastEvent = appender.getLastLogEvent()
    if (!lastEvent) {
      throw new Error("Appender did not return a LogEvent for getLastLogEvent()")
    }
    LoggerImpl._instance._criticalEvents.push(lastEvent)
    if (LoggerImpl._instance._action === LoggerImpl.ACTION.EXIT) {
      const LAST_MSG = AbstractAppender.getLayout().format(lastEvent)
      throw new ScriptError(LAST_MSG)
    }
  }
}

  /** Returns the corresponding string label for the level. */
  private static getLevelLabel(): string {
    return LoggerImpl.LEVEL_LABELS[LoggerImpl._instance._level];
  }

  /** Returns the corresponding string label for the action. */
  private static getActionLabel(): string {
    return LoggerImpl.ACTION_LABELS[LoggerImpl._instance._action];
  }

  /* Enforces instantiation lazily. If the user didn't invoke getInstance(), provides a logger
   * with default configuration. It also sends a trace event indicating the lazy initialization */
  private static initIfNeeded(): void {
    if (!LoggerImpl._instance) {
      LoggerImpl._instance = LoggerImpl.getInstance();
      const LEVEL_LABEL = `Logger.LEVEL.${LoggerImpl.getLevelLabel()}`;
      const ACTION_LABEL = `Logger.ACTION.${LoggerImpl.getActionLabel()}`;
      const MSG = `Logger instantiated via Lazy initialization with default parameters (level=${LEVEL_LABEL}, action=${ACTION_LABEL})`;
      LoggerImpl._instance.trace(MSG);
    }
  }

  // Common safeguard method, where calling initIfNeeded doesn't make sense
  private static validateInstance() {
    if (!LoggerImpl._instance) {
      const MSG = `In '${LoggerImpl.name}' class a singleton instance can't be undefined or null. Please invoke getInstance first.`;
      throw new ScriptError(MSG);
    }
  }

  /* Checks level has one of the valid values. It is required, because the way Logger.LEVEL was built,
  i.e. based on LOG_EVENT, so it doesn't check for non-valid values during compilation. That is not the
  case for Logger.ACTION. */
  private static assertValidLevel(level: typeof LoggerImpl.LEVEL[keyof typeof LoggerImpl.LEVEL]) {
    if (!Object.values(LoggerImpl.LEVEL).includes(level)) { // level not part of Logger.LEVEL
      const MSG = `The input value level='${level}', was not defined in Logger.LEVEL.`;
      throw new ScriptError(MSG);
    }
  }

  /** Validates that all appenders are of unique class types, with no null or undefined entries. */
  private static assertUniqueAppenderTypes(appenders: (Appender | null | undefined)[]): void {
    if (Utility.isEmptyArray(appenders)) {
      throw new ScriptError("Invalid input: 'appenders' must be a non-null array.");
    }

    const seen = new Set<Function>(); // ensure unique elements only
    for (const appender of appenders) {
      if (!appender) {
        throw new ScriptError("Appender list contains null or undefined entry.");
      }
      const ctor = appender.constructor;
      if (seen.has(ctor)) {
        const name = ctor.name || "UnknownAppender";
        throw new ScriptError(`Only one appender of type ${name} is allowed.`);
      }
      seen.add(ctor);
    }
  }
}


// ===================================================
// End Lightweight logging framework for Office Script
// ===================================================

// Export to globalThis for Office Scripts compatibility in Node/ts-node
if (typeof globalThis !== "undefined") {
  if (typeof LOG_EVENT !== "undefined") {
    // @ts-ignore
    globalThis.LOG_EVENT = LOG_EVENT;
  }
  
  if (typeof LogEventImpl !== "undefined") {
    // @ts-ignore
    globalThis.LogEventImpl = LogEventImpl;
  }
  if (typeof LayoutImpl !== "undefined") {
    // @ts-ignore
    globalThis.LayoutImpl = LayoutImpl;
  }
  if (typeof AbstractAppender !== "undefined") {
    // @ts-ignore
    globalThis.AbstractAppender = AbstractAppender;
  }

  if (typeof ConsoleAppender !== "undefined") {
    // @ts-ignore
    globalThis.ConsoleAppender = ConsoleAppender;
  }
  if (typeof ExcelAppender !== "undefined") {
    // @ts-ignore
    globalThis.ExcelAppender = ExcelAppender;
  }
  if (typeof ScriptError !== "undefined") {
    // @ts-ignore
    globalThis.ScriptError = ScriptError;
  }
  if (typeof LoggerImpl !== "undefined") {
    // @ts-ignore
    globalThis.LoggerImpl = LoggerImpl;
  }
  if (typeof Utility !== "undefined") {
    // @ts-ignore
    globalThis.Utility = Utility;
  }

}