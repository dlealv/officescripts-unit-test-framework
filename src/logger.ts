
// #region logger.ts
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
 * version 2.1.0
 * creation date: 2024-10-01
 */

// Enum DEFINITIONS
// --------------------

// #region enum and types
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
type LogEventFactory = (message: string, eventType: LOG_EVENT, extraFields?: LogEventExtraFields) => LogEvent
type LogEventExtraFields = {
  [key: string]: string | number | Date | (() => string)
}
type LayoutFormatter = (event: LogEvent) => string

// #endregion enum and types


// INTERFACES
// --------------------

// #region INTERFACES

/**
 * Interface for all log events to be sent to appenders.
 * Defines the structure of a log event and is intended to be immutable.
 *
 * @remarks
 * - The layout (see getLayout/setLayout in concrete implementations) is used to configure
 *   the core content formatting of the log event output. All appenders (listeners) will output
 *   the same formatted log content, ensuring consistency.
 *   Appenders may apply their own additional presentation (e.g., colors or styles) without altering
 *   the event message itself.
 * - getLayout/setLayout are static methods, from AbstractAppender class, share among
 *   all subclasses.
 * - Implementations of this interface (including user extensions) must ensure all properties
 *   are assigned during construction and remain immutable.
 * - Framework code may validate any object passed as a LogEvent to ensure all invariants hold.
 *   If you implement this interface directly, you are responsible for upholding these invariants.
 */
interface LogEvent {
  /**
   * The event type from the LOG_EVENT enum.
   * This field is immutable and must be set at construction.
   */
  readonly type: LOG_EVENT

  /**
   * The log message to be sent.
   * This field is immutable. It must not be null, undefined, or an empty string.
   */
  readonly message: string

  /**
   * The timestamp when the event was created.
   * This field is immutable and must be a valid Date instance.
   */
  readonly timestamp: Date

  /**
   * Additional metadata for the log event, for extension and contextual purposes.
   * This field is immutable and must be a plain object.
   * Intended for extensibility—avoid storing sensitive or large data here.
   */
  readonly extraFields: LogEventExtraFields

  /**
   * Returns a string representation of the log event in a human-readable, single-line format,
   * including all relevant fields.
   */
  toString(): string
}

/**
 * Interface to handle formatting of log events sent to appenders.
 * The format defines the core structure of the log message content before it is sent to appenders.
 * It is **not intended for adornment or presentation** (such as color or Excel formatting),
 * but strictly for the canonical, consistent string representation of the log event.
 *
 * @remarks
 * - Implementations must provide consistent, stateless formatting for all log events.
 * - Layout should be deterministic and MUST NOT mutate the event.
 * - Typical implementations may provide static/shared instances for consistency.
 * - Layouts are intended for core message structure, not for display/presentation logic.
 */
interface Layout {
  /**
   * Formats the given log event into its core string representation.
   * @param event - The log event to format (must be a valid, immutable LogEvent).
   * @returns The formatted string representing the event's core content.
   */
  format(event: LogEvent): string;

  /**
   * Returns a string describing the layout, ideally including the formatter function name or configuration.
   * Used for diagnostics or debugging.
   */
  toString(): string;
}

/**
 * Interface for all appenders (log destinations).
 *
 * An appender delivers log events to a specific output channel (e.g., console, Excel, file, remote service).
 *
 * @remarks
 * - Implementations must provide both:
 *   - Structured logging via `log(event: LogEvent)`
 *   - Convenience logging via `log(msg: string, event: LOG_EVENT)`
 * - Appenders are responsible for sending log events, not formatting (core formatting is handled by the Layout).
 * - Implementations should be stateless or minimize instance state except for tracking the last sent event.
 * - Common formatting and event creation concerns (such as layout or logEventFactory) are handled by the AbstractAppender base class and are not part of the interface contract.
 * - `getLastLogEvent()` is primarily for diagnostics, testing, or error reporting.
 * - `toString()` must return diagnostic information about the appender and its last log event.
 * - Provides to log methods sending a LogEvent object or a message with an event type and optional extra fields.
 *   This allows flexibility in how log events are created and sent.
 */
interface Appender {
  /**
   * Sends a structured log event to the appender.
   * @param event - The log event object to deliver.
   * @throws ScriptError if the event is invalid or cannot be delivered.
   */
  log(event: LogEvent): void

  /**
   * Sends a log message to the appender, specifying the event type and optional structured extra fields.
   * @param msg - The message to log.
   * @param type - The type of log event (from LOG_EVENT enum).
   * @param extraFields - Optional structured data (object) to attach to the log event (e.g., context info, tags).
   */
  log(msg: string, type: LOG_EVENT, extraFields?: object): void

  /**
   * Returns the last LogEvent delivered to the appender, or null if none sent yet.
   * @returns The last LogEvent object delivered, or null.
   * @throws ScriptError if the appender instance is unavailable.
   */
  getLastLogEvent(): LogEvent | null

  /**
   * Returns a string summary of the appender's state, typically including its type and last event.
   * @returns A string describing the appender and its most recent activity.
   * @throws ScriptError if the appender instance is unavailable.
   */
  toString(): string
}

/**
 * Represents a logging interface for capturing and managing log events at various levels.
 * Provides methods for logging messages, querying log state, managing appenders, and exporting logger state.
 * Implementations should ensure they should not maintain global mutable state outside the singleton and efficient 
 * log event handling.
 */
interface Logger {
  /**
   * Sends an error log event with the provided message and optional extraFields to all appenders.
   * @param msg - The error message to log.
   * @param extraFields - Optional structured data (object) to attach to the log event. May include metadata, context, etc.
   * @throws ScriptError if 
   *          - The singleton instance is not available (not instantiated)
   *          - The logger is configured to exit on critical events.
   */
  error(msg: string, extraFields?: object): void

  /**
   * Sends a warning log event with the provided message and optional extraFields to all appenders.
   * @param msg - The warning message to log.
   * @param extraFields - Optional structured data (object) to attach to the log event.
   * @throws ScriptError if 
   *          - The singleton instance is not available (not instantiated)
   *          - The logger is configured to exit on critical events.
   */
  warn(msg: string, extraFields?: object): void

  /**
   * Sends an informational log event with the provided message and optional extraFields to all appenders.
   * @param msg - The informational message to log.
   * @param extraFields - Optional structured data (object) to attach to the log event.
   * @throws ScriptError if the singleton instance is not available (not instantiated).
   */
  info(msg: string, extraFields?: object): void

  /**
   * Sends a trace log event with the provided message and optional extraFields to all appenders.
   * @param msg - The trace message to log.
   * @param extraFields - Optional structured data (object) to attach to the log event.
   * @throws ScriptError if the singleton instance is not available (not instantiated).
   */
  trace(msg: string, extraFields?: object): void

  /**
   * Gets an array of all error and warning log events sent.
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
   * @returns The log level value (e.g., OFF, ERROR, WARN, INFO, TRACE). It refers to the level of verbosity to show
   * during the logging process.
   * @throws ScriptError if the singleton instance is not available (not instantiated).
   */
  getLevel(): number

  /**
   * Gets the array of appenders currently registered with the logger.
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
  reset(): void

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

// #endregion INTERFACES

// CLASSES
// --------------------


// #region ScriptError
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
    return `${this.constructor.name}: ${this.message}\nStack trace:\n${stack}`
  }
}
// #endregion ScriptError


/**
 * Utility class providing static helper methods for logging operations.
 */
class Utility {
  /**Helpder to format the local date as a string. Ouptut in standard format: YYYY-MM-DD HH:mm:ss,SSS
  */
  public static date2Str(date: Date): string {
    // Defensive: handle null, undefined, or non-Date input
    if (!(date instanceof Date) || isNaN(date.getTime())) {
      const PREFIX = `[${Utility.name}.date2Str]: `
      return `${PREFIX}Invalid Date`
    }
    const pad = (n: number, width = 2) => n.toString().padStart(width, '0')
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)
      }-${pad(date.getDate())
      } ${pad(date.getHours())
      }:${pad(date.getMinutes())
      }:${pad(date.getSeconds())
      },${pad(date.getMilliseconds(), 3)
      }`;
  }

  /** Helper method to check for an empty array. */
  public static isEmptyArray<T>(arr: T[]): boolean {
    return (!Array.isArray(arr) || !arr.length) ? true : false;
  }

  /**
   * Validates a log event factory is a function.
   * @param factory The factory function to validate.
   * @param funName Used to identify the function name in the error message.
   * @param context 
   * @throws ScriptError if the log event factory is not a function.
   */
  public static validateLogEventFactory(
    factory: unknown, // or Function, or your specific function type
    funName?: string,
    context?: string
  ): void {
    const PREFIX = context ? `[${context}]: ` : '';
    if (typeof factory !== 'function') {
      throw new ScriptError(`${PREFIX}Invalid ${funName || "<anonymous>"}: Not a function`);
    }
  }

}


// #region LogEventImpl
/**
 * Implements the LogEvent interface, providing a concrete representation of a log event.
 * It includes properties for the event type, message, and timestamp, along with methods to manage
 * the layout used for formatting log events before sending them to appenders.
 */
class LogEventImpl implements LogEvent {
  private readonly _type: LOG_EVENT
  private readonly _message: string
  private readonly _timestamp: Date
  // Accept any additional fields
  private readonly _extraFields: LogEventExtraFields = {}
  /** Reserved keys that should not be included in the extraFields object.*/
  private static readonly RESERVED_KEYS = ['type', 'message', 'timestamp', 'toString']


  /**
   * Constructs a new LogEventImpl instance.
   * Validates the input parameters to ensure they conform to expected types and constraints.
   * @param type - The type of the log event (from LOG_EVENT enum).
   * @param message - The message to log.
   * @param timestamp - (Optional) The timestamp of the event, defaults to current time.
   * @param extraFields - (Optional) Additional fields for the log event, can include strings, numbers, dates, or functions.
   * @throws ScriptError if validation fails.
   */
  constructor(message: string, type: LOG_EVENT, extraFields?: LogEventExtraFields, timestamp: Date = new Date(),
  ) {
    LogEventImpl.validateLogEventAttrs({ type: type, message, timestamp }, extraFields, "LogEventImpl.constructor")
    this._type = type
    this._message = message
    this._timestamp = timestamp
    if (extraFields) {
      for (const [k, v] of Object.entries(extraFields)) {
        if (!LogEventImpl.RESERVED_KEYS.includes(k)) {
          this._extraFields[k] = v
        }
      }
      Object.freeze(this._extraFields)
    }
  }

  /**
   * @returns The event type from LOG_EVENT enum (immutable).
   * @override
   */
  public get type(): LOG_EVENT { return this._type }

  /**
   * @returns The message of the log event (immutable).
   * @override
   */
  public get message(): string { return this._message }

  /**
   * @returns The timestamp of the log event (immutable).
   * @override
   */
  public get timestamp(): Date { return this._timestamp }

  /**
   * Gets the extra fields of the log event.
   * @returns Returns a shallow copy of custom fields for this event. These are immutable (Object.freeze),
   * but if you allow object values in the future, document that deep mutation is not prevented.
   * @override
   */
  public get extraFields(): Readonly<LogEventExtraFields> {
    return { ...this._extraFields }
  }

  /**
   * Validates if the input object conforms to the LogEvent interface (for any implementation).
   * @throws ScriptError if event is invalid.
   */
  public static validateLogEvent(event: unknown, context?: string): void {
    const PREFIX = context ? `[${context}]: ` : `[${LogEventImpl.name}.validateLogEvent]: `
    if (typeof event !== 'object' || event == null) {
      throw new ScriptError(`${PREFIX}LogEvent must be a non-null object.`)
    }
    const e = event as { type?: unknown; message?: unknown; timestamp?: unknown; extraFields?: unknown; }
    // Validate extraFields only if present
    if ((e.extraFields !== undefined) &&
      (typeof e.extraFields !== 'object' || e.extraFields == null || Array.isArray(e.extraFields))
    ) {
      throw new ScriptError(`${PREFIX}extraFields must be a non-null plain object.`)
    }
    const CTXT = context ? context : `${LogEventImpl.name}.validateLogEvent`
    LogEventImpl.validateLogEventAttrs({
      type: e.type,
      message: e.message,
      timestamp: e.timestamp
    }, e.extraFields, CTXT) // Validate the attributes
  }

  /**
   * @returns A string representation of the log event in stardard toString format
   * @override
   */
  public toString(): string {
    const sDATE = Utility.date2Str(this._timestamp) //Local date as string
    // Get the string representation of the type, don't use LogEventImpl.eventTypeToLabel(this.type) to avoid unnecesary validation
    const sTYPE = LOG_EVENT[this.type]
    const BASE = `${this.constructor.name}: {timestamp="${sDATE}", type="${sTYPE}", message="${this._message}"`
    const HAS_EXTRA = Object.keys(this._extraFields).length > 0
    const EXTRA = HAS_EXTRA ? `, extraFields=${JSON.stringify(this.extraFields)}` : ''
    return `${BASE}${EXTRA}}`
  }

  /**
   * Returns a standardized label for the given log event.
   * @param type - The event type from 'LOG_EVENT' enum.
   * @returns A string label, e.g., '[INFO]', '[ERROR]'.
   */
  public static eventTypeToLabel(type: LOG_EVENT): string {
    const event = { type, message: "dummy", timestamp: new Date(), extraFields: {} } as LogEvent // Create a dummy message to pass the message validation
    LogEventImpl.validateLogEvent(event, "LogEventImpl.eventTypeToLabel")
    return `${LOG_EVENT[type]}`
  }

  /**
   * Validates the raw attributes for a log event, including extraFields if provided.
   * @param attrs An object containing the core attributes: type, message, timestamp.
   * @param extraFields Optional object containing additional fields to validate.
   * @param context Optional string for error context prefixing.
   * @throws ScriptError if any of the attributes are not valid.
   */
  private static validateLogEventAttrs(
    attrs: { type: unknown, message: unknown, timestamp: unknown },
    extraFields?: unknown,
    context?: string
  ): void {
    const PREFIX = context ? `[${context}]: ` : `[${LogEventImpl.name}.validateLogEventAttrs]: `;

    // Validate type
    if (typeof attrs.type !== 'number') {
      throw new ScriptError(`${PREFIX}LogEvent.type='${attrs.type}' property must be a number (LOG_EVENT enum value).`);
    }
    if (!Object.values(LOG_EVENT).includes(attrs.type as LOG_EVENT)) {
      throw new ScriptError(`${PREFIX}LogEvent.type='${attrs.type}' property is not defined in the LOG_EVENT enum.`);
    }

    // Validate message
    if (typeof attrs.message !== 'string') {
      throw new ScriptError(`${PREFIX}LogEvent.message='${attrs.message}' property must be a string.`);
    }
    if (attrs.message.trim().length === 0) {
      throw new ScriptError(`${PREFIX}LogEvent.message cannot be empty.`);
    }

    // Validate timestamp
    if (!(attrs.timestamp instanceof Date)) {
      throw new ScriptError(`${PREFIX}LogEvent.timestamp='${attrs.timestamp}' property must be a Date.`);
    }

    // Validate extraFields if provided
    if (extraFields !== undefined) {
      if (typeof extraFields !== "object" || extraFields === null || Array.isArray(extraFields)) {
        throw new ScriptError(`${PREFIX}extraFields must be a plain object.`);
      }
      for (const [k, v] of Object.entries(extraFields)) {
        if (v === undefined) {
          throw new ScriptError(`${PREFIX}extraFields[${k}] must not be undefined.`);
        }
        if (typeof v !== "string" && typeof v !== "number" &&
          !(v instanceof Date) && typeof v !== "function") {
          throw new ScriptError(`${PREFIX}extraFields[${k}] has invalid type: ${typeof v}. Must be string, number, Date, or function.`);
        }
      }
    }
  }

}
// #endregion LogEventImpl


// #region LayoutImpl
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

  /**Convninience static property to define a short formatter.
   * @see LayoutImpl.shortFormatterFun
  */
  public static shortFormatterFun: LayoutFormatter

  /**Convninient static property to define a long formatter used as default.
   * @see LayoutImpl.defaultFormatterFun
   */
  public static defaultFormatterFun: LayoutFormatter

  /**
   * Function used to convert a LogEvent into a string.
   * Set at construction time; defaults to a simple "[type] message" format.
   */
  private readonly _formatter: LayoutFormatter

  /**
   * Constructs a new LayoutImpl.
   * 
   * @param formatter - Optional. A function that formats a LogEvent as a string.
   *                    If not provided, a default formatter is used: "[timestamp] [type] message".
   *                    The formatter function synchronous and must accept a single LogEvent 
   *                    parameter and return a string.  
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
  constructor(formatter?: LayoutFormatter) {
    this._formatter = formatter ?? LayoutImpl.defaultFormatterFun
    LayoutImpl.validateLayout(this as LayoutImpl, "LayoutImpl.constructor")
  }

  /**
   * Returns the current formatter function.
   * @returns The formatter function.
   */
  public getFormatter(): LayoutFormatter {
    return this._formatter
  }

  /**
   * Formats the given log event as a string.
   * @param event - The event to format.
   * @returns A string representation of the log event.
   * @throws ScriptError if the event does not conform to the LogEvent interface.
   * @override
   */
  public format(event: LogEvent): string {
    LogEventImpl.validateLogEvent(event, "LayoutImpl.format")
    return this._formatter(event)
  }

  /**
   * @returns A string representation of the layout.
   *          If the formatter is a function, it returns the name of the function.
   * @override
   */
  public toString(): string {
    const formatterName = this._formatter?.name || "anonymous"
    return `${this.constructor.name}: {formatter: [Function: "${formatterName}"]}`
  }

  /**
   * Asserts that the provided object implements the Layout interface.
   * Checks for the public 'format' method (should be a function taking one argument).
   * Also validates the internal '_formatter' property if present, by calling validateFormatter.
   * Used by appenders to validate layout objects at runtime.
   *
   * @param layout - The object to validate as a Layout implementation
   * @param context - (Optional) Additional context for error messages
   * @throws ScriptError if:
   *    - layout is null or undefined
   *    - format is not a function or doesn't have arity 1
   *    - _formatter is present and is missing, not a function, or doesn't have arity 1
   */
  static validateLayout(layout: Layout, context?: string) {
    const PREFIX = context ? `[${context}]: ` : "[LayoutImpl.validateLayout]: "
    if (!layout || typeof layout !== "object") {
      throw new ScriptError(PREFIX + "Invalid Layout: layout object is null or undefined")
    }
    if (typeof layout.format !== "function" || layout.format.length !== 1) {
      throw new ScriptError(
        `{PREFIX} Invalid Layout: The 'format' method must be a function accepting a single LogEvent argument. ` +
        `See LayoutImpl documentation for usage.`
      );
    }
    if (layout instanceof LayoutImpl) {
      LayoutImpl.validateFormatter(layout._formatter, context)
    }
  }

  /**
   * Validates that the provided value is a valid formatter function
   * for use in LayoutImpl (_formatter property). The formatter must be
   * a function accepting a single LogEvent argument and must return a non-empty, non-null string.
   *
   * @param formatter - The candidate formatter function to validate
   * @param context - (Optional) Additional context for error messages
   * @throws ScriptError if formatter is missing, not a function, doesn't have arity 1,
   *                     or returns null/empty string for a sample event.
   */
  static validateFormatter(formatter: LayoutFormatter, context?: string) {
    const PREFIX = context ? `[${context}]: ` : "[LayoutImpl.validateFormatter]: ";
    if (typeof formatter !== "function" || formatter.length !== 1) {
      throw new ScriptError(
        PREFIX +
        "Invalid Layout: The internal '_formatter' property must be a function accepting a single LogEvent argument. See LayoutImpl documentation for usage."
      );
    }
    // Try calling with a mock event
    const mockEvent: LogEvent = {
      type: LOG_EVENT.INFO,
      message: "test",
      timestamp: new Date(),
      extraFields: {},
    };
    const result = formatter(mockEvent);
    if (
      typeof result !== "string" ||
      result === "" ||
      result == null
    ) {
      throw new ScriptError(
        PREFIX +
        "Formatter function must return a non-empty string for a valid LogEvent. Got: " +
        (result === "" ? "empty string" : String(result))
      );
    }
  }

}

// Assign the static formatters outside the class
/**
 * Convenience public constant to help users to define a short format for log events. 
 * Formats a log event as a short string as follows '[type] message'.
 * If extraFields are present in the event, they will be appended as a JSON object (surrounded by braces) to the output.
 * Example: [ERROR] Something bad happened {"user":"dlealv","id":42}
 * Defined as a named function to ensure toString() returns the function name.
 */

LayoutImpl.shortFormatterFun = Object.freeze(function shortLayoutFormatterFun(event: LogEvent): string {
  const sType = LOG_EVENT[event.type]
  let extraFieldsStr = ""
  if (event.extraFields && Object.keys(event.extraFields).length > 0) {
    extraFieldsStr = ` ${JSON.stringify(event.extraFields)}` // JSON.stringify includes the braces
  }
  return `[${sType}] ${event.message}${extraFieldsStr}`
})

/**
 * Default formatter function. Created as a named function. Formats a log event as [timestamp] [type] message.
 * The timestamp is formatted as YYYY-MM-DD HH:mm:ss,SSS.
 * If extraFields are present in the event, they will be appended as a JSON object (surrounded by braces) to the  output.
 * Example: [2025-06-19 15:06:41,123] [ERROR] Something bad happened {"user":"dlealv","id":42}
 * Defined as a named function to ensure toString() returns the function name.
 */

LayoutImpl.defaultFormatterFun = Object.freeze(function defaultLayoutFormatterFun(event: LogEvent): string {
  const sDATE = Utility.date2Str(event.timestamp)
  const sType = LOG_EVENT[event.type]
  let extraFieldsStr = ""
  if (event.extraFields && Object.keys(event.extraFields).length > 0) {
    extraFieldsStr = ` ${JSON.stringify(event.extraFields)}` // JSON.stringify includes the braces
  }
  return `[${sDATE}] [${sType}] ${event.message}${extraFieldsStr}`
})

// #endregion LayoutImpl


// #region AbstractAppender
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
  // Default factory function to create LogEvent instances, if was not set before.
  public static defaultLogEventFactoryFun: LogEventFactory
  // Static layout shared by all events
  private static _layout: Layout | null = null
  // Static log event factory function used to create LogEvent instances.
  private static _logEventFactory: LogEventFactory | null = null
  private _lastLogEvent: LogEvent | null = null // The last event sent by the appender

  /**
   * Constructs a new AbstractAppender instance. Nothing is initialized, because the class only has static properties
   * that are lazy initialized or set by the user.
   */
  protected constructor() { }

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
   * Sets the log event factory function used to create LogEvent instances if it was not set before.
   * @param logEventFactory - A factory function to create LogEvent instances.
   *                          Must have the signature (message: string, eventType: LOG_EVENT) => LogEvent.
   *                          If not provided, a default factory function is used.
   * @throws ScriptError if the log event factory is not a valid function with the expected signature.
   * @example
   * // Example: Custom LogEvent to be used to specify the environment where the log event was created.
   * let prodLogEventFactory: LogEventFactory
      = function prodLogEventFactoryFun(message: string, eventType: LOG_EVENT) {
        return new LogEventImpl("PROD-" + message, eventType) // add environment prefix
      }
   * AbstractAppender.setLogEventFactory(prodLogEventFactory) // Now all appenders will use ProdLogEvent
   */
  public static setLogEventFactory(logEventFactory: LogEventFactory): void {
    if (!AbstractAppender._logEventFactory) {
      AbstractAppender.validateLogEventFactory(logEventFactory, "logEventFactory", "AbstractAppender.setLogEventFactory")
      AbstractAppender._logEventFactory = logEventFactory
    }
  }

  /** Gets the log event factory function used to create LogEvent instances. If it was not set before, it returns the default factory function.
   * @returns The log event factory function.
   */
  public static getLogEventFactory(): LogEventFactory {
    if (!AbstractAppender._logEventFactory) {
      AbstractAppender._logEventFactory = AbstractAppender.defaultLogEventFactoryFun // Default factory if not set
    }
    return AbstractAppender._logEventFactory
  }

  /**
   * Sets the layout associated to all events, the layout is assigned only if it was not set before.
   * @param layout - The layout to set.
   * @throws ScriptError if the layout is not a valid Layout implementation.
   */
  public static setLayout(layout: Layout): void {
    const CONTEXT = "AbstractAppender.setLayout"
    if (!AbstractAppender._layout) {
      LayoutImpl.validateLayout(layout, CONTEXT)
      AbstractAppender._layout = layout
    }
  }

  /**
   * Log a message or event.
   * @param arg1 - LogEvent or message string.
   * @param arg2 - LOG_EVENT, only required if arg1 is a string.
   * @param arg3 - extraFields, only used if arg1 is a string.
   * @override
   */

  public log(arg1: LogEvent | string, arg2?: LOG_EVENT, arg3?: LogEventExtraFields): void {
    const CONTEXT = `AbstractAppender.log`
    const PREFIX = `[${CONTEXT}]: `
    if (typeof arg1 === "string") {
      if (arg2 === undefined || !(Object.values(LOG_EVENT) as unknown[]).includes(arg2)) {
        throw new ScriptError(`${PREFIX}event type='${arg2}' must be provided and must be a valid LOG_EVENT value.`)
      }
      const event = AbstractAppender.getLogEventFactory()(arg1, arg2, arg3)
      this.sendEvent(event, CONTEXT)
    } else {
      LogEventImpl.validateLogEvent(arg1, CONTEXT)
      this.sendEvent(arg1, CONTEXT)
    }
  }

  /**
   * @returns The last log event sent by the appender.
   * @override
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

  // #TEST-ONLY-START
  /** Sets to null the log event factory, useful for running different scenarios.*/
  public static clearLogEventFactory(): void {
    AbstractAppender._logEventFactory = null // Force full re-init
  }
  // #TEST-ONLY-END

  /**
   * Returns a string representation of the appender.
   * It includes the information from the base class plus the information of the current class, 
   * so far this class doesn't have additional properties to show.
   * @returns A string representation of the appender.
   * @override
   */
  public toString(): string {
    const NAME = "AbstractAppender" // since it can be extended, we use the class name as literal
    const LAYOUT = AbstractAppender._layout
    const LAYOUT_STR = LAYOUT ? LAYOUT.toString() : "null"
    const FACTORY_STR = AbstractAppender._logEventFactory ? AbstractAppender._logEventFactory.name || "anonymous" : "null"
    const LAST_LOG_EVENT_STR = this._lastLogEvent ? this._lastLogEvent.toString() : "null"
    return `${NAME}: {layout=${LAYOUT_STR}, logEventFactory="${FACTORY_STR}", lastLogEvent=${LAST_LOG_EVENT_STR}}`
  }

  /**
   * Send the event to the appropriate destination. The event is stored based on the 
   * @param event - The log event to be sent.
   * @param context - (Optional) A string to provide additional context in case of an error.
   * @throws ScriptError if 
   *         - The event is not a valid LogEvent.
   */
  protected abstract sendEvent(event: LogEvent, context?: string): void

  /**
 * Send the event to the appropriate destination.
 * @param event - The log event to be sent.
 * @throws ScriptError if 
 *         - The event is not a valid LogEvent.
 * @remarks
 * Subclasses **must** call `setLastLogEvent(event)` after successfully sending the event,
 * otherwise `getLastLogEvent()` will not reflect the most recent log event.
 */
  protected setLastLogEvent(event: LogEvent): void {
    const CONTEXT = "AbstractAppender.setLastLogEvent"
    LogEventImpl.validateLogEvent(event, CONTEXT) // Validate the event
    this._lastLogEvent = event // Set the last log event
  }

  // #TEST-ONLY-START
  /**To ensure when invoking clearInstance in a sub-class it also clear the last log event */
  protected clearLastLogEvent(): void {
    this._lastLogEvent = null
  }
  // #TEST-ONLY-END

  /**
   * Validates a log event factory is a function.
   * @param factory The factory function to validate.
   * @param funName Used to identify the function name in the error message.
   * @param context 
   * @throws ScriptError if the log event factory is not a function.
   */
  private static validateLogEventFactory(
    factory: unknown, // or Function, or your specific function type
    funName?: string,
    context?: string
  ): void {
    const PREFIX = context ? `[${context}]: ` : '';
    if (typeof factory !== 'function') {
      throw new ScriptError(`${PREFIX}Invalid ${funName || "<anonymous>"}: Not a function`);
    }
  }

}

// Functions outside of the class:
AbstractAppender.defaultLogEventFactoryFun = Object.freeze(
  function defaultLogEventFactoryFun(message: string, eventType: LOG_EVENT, extraFields?: LogEventExtraFields) {
    return new LogEventImpl(message, eventType, extraFields)
  }
)

// #endregion AbstractAppender


// #region ConsoleAppender
/**
 * Singleton appender that logs messages to the Office Script console. It is used as default appender, 
 * if no other appender is defined. The content of the message event sent can be customized via
 * any Layout implementation, but by default it uses the LayoutImpl.
 * Usage:
 * - Call ConsoleAppender.getInstance() to get the appender
 * - Automatically used if no other appender is defined
 * Warning: The console appender is a singleton, so it should not be instantiated multiple times.
 * @example:
 * // Add console appender to the Logger
 * Logger.addAppender(ConsoleAppender.getInstance())
 */
class ConsoleAppender extends AbstractAppender implements Appender {
  private static _instance: ConsoleAppender | null = null // Instance of the singleton pattern

  /**
   * Private constructor to prevent user instantiation.
   * Initializes the appender without a layout, using the default layout if not set (lazy initialization).
   */
  private constructor() {
    super() // Call the parent constructor without layout
  }

  /**Override toString method. Show the last message event sent
   * @throws ScriptError If the singleton was not instantiated
   */
  public toString(): string {
    ConsoleAppender.validateInstance("ConsoleAppender.toString") // Validate the singleton instance
    const name = this.constructor.name
    return `${super.toString()} ${name}: {}`
  }

  /** Gets the singleton instance of the class. 
   * @returns The singleton instance of the class. If it was not instantiated, it creates
   *          a new instance.
   */
  public static getInstance(): ConsoleAppender {
    if (!ConsoleAppender._instance) {
      ConsoleAppender._instance = new ConsoleAppender()
    }
    return ConsoleAppender._instance
  }

  // #TEST-ONLY-START
  /** Sets to null the singleton instance, useful for running different scenarios.
   * It also sets to null the parent property _lastLogEvent, so the last log event is cleared.
   * @remarks Mainly intended for testing purposes. The state of the singleton will be lost.
   *          This method only exist in src folder it wont be deployed in dist folder (production).  
   * @example There is no way to empty the last message sent after the instance was created unless
   * you reset it.
   * appender:ConsoleAppender = ConsoleAppender.getInstance()
   * appender = ConsoleAppender.log("info event", LOG_EVENT.INFO)
   * appender.getLastLogEvent().message         // Output: info event"
   * ConsoleAppender.clearInstance()            // clear the singleton
   * appender = ConsoleAppender.getInstance()   // restart the singleton
   * appender.getLastLogEvent().message         // Output: ""
   */
  public static clearInstance(): void {
    if (ConsoleAppender._instance) {
      ConsoleAppender._instance.clearLastLogEvent()
    }
    ConsoleAppender._instance = null
  }
  // #TEST-ONLY-END

  /**
   * Internal method to send the event to the console.
   * @param event - The log event to output.
   * @param context - (Optional) A string to provide additional context in case of an error.
   * @throws ScriptError if 
   *          The event is not a valid LogEvent.
   *          The instance is not available (not instantiated).
   * @override
   */
  protected sendEvent(event: LogEvent, context?: string): void {
    const CTX = context ? context : `${this.constructor.name}.sendEvent`
    LogEventImpl.validateLogEvent(event, CTX) // Validate the event
    ConsoleAppender.validateInstance(CTX) // Validate the instance
    this.setLastLogEvent(event)
    // format the output using the layout that gets lazy initialized if it was not set before
    const MSG = AbstractAppender.getLayout().format(event)
    console.log(MSG)

  }

  /** @internal
  * Common safeguard method, where calling initIfNeeded doesn't make sense.
  * @param context - (Optional) A string to provide additional context in case of an error.
  * @throws ScriptError In case the singleton was not initialized.
  */
  private static validateInstance(context?: string): void {
    if (!ConsoleAppender._instance) {
      const PREFIX = context ? `[${context}]: ` : `[${ConsoleAppender.name}.validateInstance]: `
      const MSG = `${PREFIX}A singleton instance can't be undefined or null. Please invoke getInstance first.`
      throw new ScriptError(MSG)
    }
  }

}

// #endregion ConsoleAppender


// #region ExcelAppender
/**
 * Singleton appender that logs messages to a specified Excel cell.
 * Logs messages in color based on the LOG_EVENT enum:
 * - ERROR: red, WARN: orange, INFO: green, TRACE: gray (defaults can be customized)
 * Usage:
 * - Must call ExcelAppender.getInstance(range) once with a valid single cell range
 * - range is used to display log messages
 * Warning: The Excel appender is a singleton, so it should not be instantiated multiple times.
 * @example:
 * ```ts
 * const range = workbook.getWorksheet("Log").getRange("A1")
 * Logger.addAppender(ExcelAppender.getInstance(range)) // Add appender to the Logger
 * ```
*/
class ExcelAppender extends AbstractAppender implements Appender {
  /**
   * Default colors for log events, used if no custom colors are provided.
   * These colors are defined as hex strings (without the # prefix).
   * The colors can be customized by passing a map of LOG_EVENT types to hex color strings
   * when calling getInstance(). Default colors are:
   * - ERROR: "9c0006" (red)
   * - WARN: "ed7d31" (orange)
   * - INFO: "548235" (green)
   * - TRACE: "7f7f7f" (gray)
   */
  public static readonly DEFAULT_EVENT_FONTS = Object.freeze({
    [LOG_EVENT.ERROR]: "9c0006",  // RED
    [LOG_EVENT.WARN]: "ed7d31",   // ORANGE
    [LOG_EVENT.INFO]: "548235",   // GREEN
    [LOG_EVENT.TRACE]: "7f7f7f"   // GRAY
  } as const);

  /**
   * Instance-level font map for current appender configuration.
   * Maps LOG_EVENT types to hex font strings.
   */
  private readonly _eventFonts: Record<LOG_EVENT, string>

  /* Regular expression to validate hexadecimal fonts */
  private static readonly HEX_REGEX = Object.freeze(/^#?[0-9A-Fa-f]{6}$/)

  private static _instance: ExcelAppender | null = null; // Instance of the singleton pattern
  private readonly _msgCellRng: ExcelScript.Range
  /* Required for Office Script limitation, only use getAddress, the first time _msgCellRng  is assigned, then 
  use this property. Calling this._msgCellRng.getAddress() fails in toString(). The workaround is to create
  this artificial property. */
  private _msgCellRngAddress: string

  /**
   * Private constructor to prevent user invocation.
   * @remarks Office Script limitation. Cannot call ExcelScript API methods on Office objects inside a class constructor, instead
   * we do such API calls in the getInstance() method.
   */
  private constructor(msgCellRng: ExcelScript.Range, eventFonts: Record<LOG_EVENT, string> = ExcelAppender.DEFAULT_EVENT_FONTS
  ) {
    super()
    this._msgCellRng = msgCellRng
    this.clearCellIfNotEmpty() // it can't be called in the construtor due to Office Script limitations
    this._eventFonts = eventFonts
  }

  // Setters and getters for the private properties
  /**
   * Returns the map of event types to font colors used by this appender.
   * @returns A defensive copy of the event fonts map.
   * @remarks The keys are LOG_EVENT enum values, and the values are hex color strings.
   */
  public getEventFonts(): Record<LOG_EVENT, string> {
    return { ...this._eventFonts }; // Defensive copy
  }

  /**
   * Returns the Excel range where log messages are written.
   * @returns The ExcelScript.Range object representing the message cell range.
   * @remarks This is the cell where log messages will be displayed.
   */
  public getMsgCellRng(): ExcelScript.Range {
    return { ...this._msgCellRng }
  }

  /**
   * Returns the singleton ExcelAppender instance, creating it if it doesn't exist.
   * On first call, requires a valid single cell Excel range to display log messages and optional
   * color customizations for different log events (LOG_EVENT). Subsequent calls ignore parameters
   * and return the existing instance.
   * @param msgCellRng - Excel range where log messages will be written. Must be a single cell and
   * not null or undefined.
   * @param eventFonts - Optional. A map of LOG_EVENT types to hex color codes for the font colors.
   *                     If not provided, defaults to the predefined colors in DEFAULT_EVENT_FONTS.
   *                     The user can provide just the colors they want to customize,
   *                     the rest will use the default colors.
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
   * @see DEFAULT_EVENT_FONTS
  */
  public static getInstance(
    msgCellRng: ExcelScript.Range, eventFonts: Record<LOG_EVENT, string> = ExcelAppender.DEFAULT_EVENT_FONTS
  ): ExcelAppender {
    const PREFIX = `[${ExcelAppender.name}.getInstance]: `
    if (!ExcelAppender._instance) {
      if (!msgCellRng || !msgCellRng.setValue) {
        const MSG = `${PREFIX}A valid ExcelScript.Range for input argument msgCellRng is required.`;
        throw new ScriptError(MSG)
      }
      if (msgCellRng.getCellCount() != 1) {
        const MSG = `${PREFIX}Input argument msgCellRng must represent a single Excel cell.`;
        throw new ScriptError(MSG)
      }
      // Enhanced Excel Range check in getInstance:
      if (!msgCellRng || typeof msgCellRng.setValue !== "function" ||
        typeof msgCellRng.getValue !== "function" ||
        typeof msgCellRng.getFormat !== "function" ||
        typeof msgCellRng.clear !== "function") {
        const MSG = `${PREFIX}A valid ExcelScript.Range for input argument msgCellRng is required.`
        throw new ScriptError(MSG)
      }
      // Checking valid hexadecimal color
      ExcelAppender.validateLogEventMappings() // Validate all LOG_EVENT mappings for fonts
      const CONTEXT = `${ExcelAppender.name}.getInstance`;
      // Merge defaults with user-provided values (user takes precedence)
      const fonts: Record<LOG_EVENT, string> = {
        ...ExcelAppender.DEFAULT_EVENT_FONTS,
        ...(eventFonts ?? {})
      };
      for (const [event, font] of Object.entries(fonts)) {
        const label = LOG_EVENT[Number(event)];
        ExcelAppender.validateFont(font, label, CONTEXT)
      }
      ExcelAppender._instance = new ExcelAppender(msgCellRng, fonts)
      // Invoking Office Script API method, can't called in the `constructor` due to Office Script limitations
      ExcelAppender._instance.clearCellIfNotEmpty("")
      ExcelAppender._instance._msgCellRngAddress = msgCellRng.getAddress() // Store the address of the range for later use
      ExcelAppender._instance._msgCellRng.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center)
    }
    return ExcelAppender._instance
  }

  // #TEST-ONLY-START
  /**
   * Sets to null the singleton instance, useful for running different scenarios.
   * It also sets to null the parent property _lastLogEvent, so the last log event is cleared.
   * @remarks Mainly intended for testing purposes. The state of the singleton will be lost.
   *          This method only exist in src folder, it wont be deployed in dist folder (production).  
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
      ExcelAppender._instance.clearLastLogEvent();
    }
    ExcelAppender._instance = null // Clear the singleton instance
  }
  // #TEST-ONLY-END

  /**
   * Shows instance configuration plus last message sent by the appender
   * @throws ScriptError, if the singleton was not instantiated.
   */
  public toString(): string {
    ExcelAppender.validateInstance("ExcelAppender.toString")
    const NAME = this.constructor.name
    // Use enum reverse mapping for label
    const EVENT_COLORS = Object.entries(this._eventFonts).map(
      ([key, value]) =>
        `${LOG_EVENT[Number(key)]}="${value}"`
    ).join(",");
    const output = `${super.toString()} ${NAME}: {msgCellRng(address)="${this._msgCellRngAddress}", `
      + `eventfonts={${EVENT_COLORS}}}`
    return output
  }

  /**
   * Sets the value of the cell, with the event message, using the font defined for the event type,
   * if not font was defined it doesn't change the font of the cell.
   * @param event a value from enum LOG_EVENT.
   * @throws ScriptError in case event is not a valid LOG_EVENT enum value.
   * @override
   */
  protected sendEvent(event: LogEvent, context?: string): void {
    const CTX = context ? context : `${this.constructor.name}.sendEvent`
    ExcelAppender.validateInstance(CTX)
    LogEventImpl.validateLogEvent(event, CTX) // Validate the event
    const FONT = this._eventFonts[event.type] ?? null
    // If no color defined for event type, use default font color (do not throw)
    if (FONT) {
      this._msgCellRng.getFormat().getFont().setColor(FONT)
    }
    const MSG = AbstractAppender.getLayout().format(event)
    this.clearCellIfNotEmpty(MSG)
    this._msgCellRng.setValue(MSG)
    this._msgCellRng.getValue() // Explicitly access the cell to ensure it commits the update
    this.setLastLogEvent(event)
  }

  // Common safeguard method
  private static validateInstance(context?: string): void {
    const PREFIX = context ? `[${context}]: ` : `[${ExcelAppender.name}]: `
    // If the instance is not defined, throw an error
    if (!ExcelAppender._instance) {
      const MSG = `${PREFIX}A singleton instance can't be undefined or null. Please invoke getInstance first`;
      throw new ScriptError(MSG)
    }
  }

  /**
   * Validate color is a valid hexadecimal color. Normalize the color
   * by removing the leading '#' if present.
   * @param color - The color string to validate.
   * @param name - The name of the event type (e.g., "error", "warning", etc.).
   * @param context - (Optional) Additional context for error messages.
   * @throws ScriptError if the color is not a valid 6-digit hexadecimal color.
   * @remarks The color must be in 'RRGGBB' or '#RRGGBB' format.
   *          If the color is not valid, it throws a ScriptError with a message indicating the issue.
   */
  private static validateFont(color: string, name: string, context?: string): void {
    const PREFIX = context ? `[${context}]: ` : `[${ExcelAppender.name}.assertColor]: `
    if (typeof color !== "string" || !color) {
      const MSG = `${PREFIX}The input value '${color}' for '${name}' event is missing or not a string. Please provide a 6-digit hexadecimal color as 'RRGGBB' or '#RRGGBB'.`
      throw new ScriptError(MSG)
    }
    const normalized = color.startsWith("#") ? color.substring(1) : color
    if (!ExcelAppender.HEX_REGEX.test(normalized)) {
      const MSG = `${PREFIX}The input value '${color}' for '${name}' event is not a valid 6-digit hexadecimal color. Please use 'RRGGBB' or '#RRGGBB' format.`
      throw new ScriptError(MSG)
    }
  }

  /** Validates that all log events are properly mapped to colors and fonts. */
  private static validateLogEventMappings(): void {
    const logEventValues = Object.values(LOG_EVENT).filter(v => typeof v === "number") as LOG_EVENT[]
    const missingColor = logEventValues.filter(ev => !(ev in ExcelAppender.DEFAULT_EVENT_FONTS))
    if (missingColor.length > 0) {
      throw new ScriptError(
        `[ExcelAppender]: LOG_EVENT enum is not fully mapped in DEFAULT_EVENT_FONTS. Missing: color=${missingColor.map(ev => LOG_EVENT[ev]).join(", ")}`
      );
    }
  }

  /**
   * Clears the message cell only if it is not empty.
   * @remarks Defined before constructor to ensure Script Office compatibility, since it is used in the constructor.
   */
  private clearCellIfNotEmpty(nextValue?: string): void {
    const value = this._msgCellRng.getValue()
    if (value !== undefined && value !== null && value !== "" && value !== nextValue) {
      this._msgCellRng.clear(ExcelScript.ClearApplyTo.contents)
    }
  }

}
// #endregion ExcelAppender


// #region LoggerImpl
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
 * If no appender is defined when a log event occurs, LoggerImpl will automatically create and add a ConsoleAppender.
 * This ensures that log messages are not silently dropped.
 * You may replace or remove this appender at any time using setAppenders() or removeAppender().
 * @example
 * // Minimal logger usage; ConsoleAppender is auto-added if none specified
 * LoggerImpl.getInstance().info("This message will appear in the console by default.");
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
    acc[value] = key
    return acc
  }, {} as Record<string, string>)

  // Equivalent labels from ACTION
  private static readonly ACTION_LABELS = Object.entries(LoggerImpl.ACTION).reduce((acc, [key, value]) => {
    acc[value] = key;
    return acc;
  }, {} as Record<number, string>)

  // Attributes
  private static _instance: LoggerImpl | null = null; // Instance of the singleton pattern
  private static readonly DEFAULT_LEVEL = LoggerImpl.LEVEL.WARN
  private static readonly DEFAULT_ACTION = LoggerImpl.ACTION.EXIT

  private readonly _level: typeof LoggerImpl.LEVEL[keyof typeof LoggerImpl.LEVEL] = LoggerImpl.DEFAULT_LEVEL
  private readonly _action: typeof LoggerImpl.ACTION[keyof typeof LoggerImpl.ACTION] = LoggerImpl.DEFAULT_ACTION
  private _criticalEvents: LogEvent[] = []; // Collects all ERROR and WARN events only
  private _errCnt = 0;   // Counts the number of error events found
  private _warnCnt = 0;  // Counts the number of warning events found
  private _appenders: Appender[] = []; // List of appenders

  // Private constructor to prevent user instantiation
  private constructor(
    level: typeof LoggerImpl.LEVEL[keyof typeof LoggerImpl.LEVEL] = LoggerImpl.DEFAULT_LEVEL,
    action: typeof LoggerImpl.ACTION[keyof typeof LoggerImpl.ACTION] = LoggerImpl.DEFAULT_ACTION,
  ) {
    this._action = action
    this._level = level
  }

  // Getters
  /** 
   * @returns An array with error and warning event messages only.
   * @throws ScriptError If the singleton was not instantiated.
   * @override
   */

  public getCriticalEvents(): LogEvent[] {
    LoggerImpl.validateInstance("LoggerImpl.getCriticalEvents")
    return this._criticalEvents
  }

  /** 
   * @returns Total number of error message events sent to the appenders.
   * @throws ScriptError If the singleton was not instantiated.
   * @override
   */
  public getErrCnt(): number {
    LoggerImpl.validateInstance("LoggerImpl.getErrCnt") // Validate the instance
    return this._errCnt
  }

  /**
   * @returns Total number of warning events sent to the appenders.
   * @throws ScriptError If the singleton was not instantiated.
   * @override
   */
  public getWarnCnt(): number {
    LoggerImpl.validateInstance("LoggerImpl.getWarnCnt") // Validate the instance
    return this._warnCnt
  }

  /** 
   * @returns The action to take in case of errors or warning log events.
   * @throws ScriptError If the singleton was not instantiated.
   * @override
   */
  public getAction(): typeof LoggerImpl.ACTION[keyof typeof LoggerImpl.ACTION] {
    LoggerImpl.validateInstance("LoggerImpl.getAction") // Validate the instance
    return this._action
  }

  /** 
   * Returns the level of verbosity allowed in the Logger. The levels are incremental, i.e.
   * it includes all previous levels. For example: Logger.WARN includes warnings and errors since
   * Logger.ERROR is lower.
   * @returns The current log level.
   * @throws ScriptError If the singleton was not instantiated.
   * @override
   */
  public getLevel(): typeof LoggerImpl.LEVEL[keyof typeof LoggerImpl.LEVEL] {
    LoggerImpl.validateInstance("LoggerImpl.getLevel") // Validate the instance
    return this._level
  }

  /**
   * @returns Array with appenders subscribed to the Logger.
   * @throws ScriptError If the singleton was not instantiated.
   * @override
   */
  public getAppenders(): Appender[] {
    LoggerImpl.validateInstance("LoggerImpl.getAppenders") // Validate the instance
    return this._appenders
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
   * @override
   * @see addAppender
   */
  public setAppenders(appenders: Appender[]) {
    const CONTEXT = `${LoggerImpl.name}.setAppenders`
    LoggerImpl.validateInstance(CONTEXT) // Validate the instance
    LoggerImpl.assertUniqueAppenderTypes(appenders, CONTEXT)
    this._appenders = appenders
  }

  /**
   * Adds an appender to the list of appenders.
   * @param appender The appender to add.
   * @throws ScriptError If the singleton was not instantiated,
   *                     if the input argument is null or undefined,
   *                     or if it breaks the class uniqueness of the appenders.
   *                     All appenders must be from a different implementation of the Appender class.
   * @override
   * @see setAppenders
   */
  public addAppender(appender: Appender): void {
    LoggerImpl.validateInstance("LoggerImpl.addAppender") // Validate the instance
    if (!appender) { // It must be a valid appender
      const PREFIX = `[${LoggerImpl.name}.addAppender]: `
      const MSG = `${PREFIX}You can't add an appender that is null or undefined`
      throw new ScriptError(MSG)
    }
    const newAppenders = [...LoggerImpl._instance._appenders, appender]
    LoggerImpl.assertUniqueAppenderTypes(newAppenders, "LoggerImpl.addAppender") // Validate uniqueness
    this._appenders.push(appender)
  }

  /**
   * Returns the singleton Logger instance, creating it if it doesn't exist.
   * If the Logger is created during this call, the provided 'level' and 'action'
   * parameters initialize the log level and error-handling behavior.
   * @remarks Subsequent calls ignore these parameters and return the existing instance.
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
    action: typeof LoggerImpl.ACTION[keyof typeof LoggerImpl.ACTION] = LoggerImpl.DEFAULT_ACTION): LoggerImpl {
    if (!this._instance) {
      const CONTEXT = `${LoggerImpl.name}.getInstance`
      LoggerImpl.validateLevelEnumIntegrity()
      LoggerImpl.assertValidLevel(level, CONTEXT) // Validate the level
      this._instance = new LoggerImpl(level, action)
    }
    return this._instance;
  }

  /**
 * If the list of appenders is not empty, removes the appender from the list.
 * @param appender The appender to remove.
 * @throws ScriptError If the singleton was not instantiated.
 */
  public removeAppender(appender: Appender): void {
    const CONTEXT = `${LoggerImpl.name}.removeAppender`
    LoggerImpl.validateInstance(CONTEXT) // Validate the instance
    const appenders = this._appenders
    if (!Utility.isEmptyArray(appenders)) {
      const index = this._appenders.indexOf(appender)
      if (index > -1) {
        this._appenders.splice(index, 1) // Remove one element at index
      }
    }
  }

  /**
 * Sends an error log message (with optional structured extra fields) to all appenders if the level allows it.
 * The level has to be greater than or equal to Logger.LEVEL.ERROR to send this event to the appenders.
 * After the message is sent, it updates the error counter.
 * @param msg - The error message to log.
 * @param extraFields - Optional structured data to attach to the log event (e.g., context info, tags).
 * @remarks
 * If no singleton was defined, it does lazy initialization with default configuration.
 * If no appender was defined, it does lazy initialization to ConsoleAppender.
 * @throws ScriptError Only if level is greater than Logger.LEVEL.OFF and the action is Logger.ACTION.EXIT.
 */
  public error(msg: string, extraFields?: LogEventExtraFields): void {
    this.log(msg, LOG_EVENT.ERROR, extraFields)
  }

  /**
   * Sends a warning log message (with optional structured extra fields) to all appenders if the level allows it.
   * The level has to be greater than or equal to Logger.LEVEL.WARN to send this event to the appenders.
   * After the message is sent, it updates the warning counter.
   * @param msg - The warning message to log.
   * @param extraFields - Optional structured data to attach to the log event (e.g., context info, tags).
   * @remarks
   * If no singleton was defined, it does lazy initialization with default configuration.
   * If no appender was defined, it does lazy initialization to ConsoleAppender.
   * @throws ScriptError Only if level is greater than Logger.LEVEL.ERROR and the action is Logger.ACTION.EXIT.
   */
  public warn(msg: string, extraFields?: LogEventExtraFields): void {
    this.log(msg, LOG_EVENT.WARN, extraFields)
  }

  /**
   * Sends an info log message (with optional structured extra fields) to all appenders if the level allows it.
   * The level has to be greater than or equal to Logger.LEVEL.INFO to send this event to the appenders.
   * @param msg - The informational message to log.
   * @param extraFields - Optional structured data to attach to the log event (e.g., context info, tags).
   * @remarks
   * If no singleton was defined, it does lazy initialization with default configuration.
   * If no appender was defined, it does lazy initialization to ConsoleAppender.
   */
  public info(msg: string, extraFields?: LogEventExtraFields): void {
    this.log(msg, LOG_EVENT.INFO, extraFields)
  }

  /**
   * Sends a trace log message (with optional structured extra fields) to all appenders if the level allows it.
   * The level has to be greater than or equal to Logger.LEVEL.TRACE to send this event to the appenders.
   * @param msg - The trace message to log.
   * @param extraFields - Optional structured data to attach to the log event (e.g., context info, tags).
   * @remarks
   * If no singleton was defined, it does lazy initialization with default configuration.
   * If no appender was defined, it does lazy initialization to ConsoleAppender.
   */
  public trace(msg: string, extraFields?: LogEventExtraFields): void {
    this.log(msg, LOG_EVENT.TRACE, extraFields)
  }

  /**
   * @returns true if an error log event was sent to the appenders, otherwise false.
   * @throws ScriptError If the singleton was not instantiated.
   */
  public hasErrors(): boolean {
    const CONTEXT = `${LoggerImpl.name}.hasErrors`
    LoggerImpl.validateInstance(CONTEXT)
    return this._errCnt > 0
  }

  /**
   * @returns true if a warning log event was sent to the appenders, otherwise false.
   * @throws ScriptError If the singleton was not instantiated.
   */
  public hasWarnings(): boolean {
    const CONTEXT = `${LoggerImpl.name}.hasWarnings`
    LoggerImpl.validateInstance(CONTEXT)
    return this._warnCnt > 0
  }

  /**
   * @returns true if some error or warning event has been sent by the appenders, otherwise false.
   * @throws ScriptError If the singleton was not instantiated.
   */
  public hasMessages(): boolean {
    const CONTEXT = `${LoggerImpl.name}.hasMessages`
    LoggerImpl.validateInstance(CONTEXT)
    return this._criticalEvents.length > 0
  }

  /**
   * Resets the Logger history, i.e., state (errors, warnings, message summary). It doesn't reset the appenders.
   * @throws ScriptError If the singleton was not instantiated.
   */
  public reset(): void {
    const CONTEXT = `${LoggerImpl.name}.clear`
    LoggerImpl.validateInstance(CONTEXT)
    this._criticalEvents = []
    this._errCnt = 0
    this._warnCnt = 0
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
    const CONTEXT = `${LoggerImpl.name}.exportState`
    LoggerImpl.validateInstance(CONTEXT); // Validate the instance
    const levelKey = Object.keys(LoggerImpl.LEVEL).find(k => LoggerImpl.LEVEL[k as keyof typeof LoggerImpl.LEVEL] === this._level);
    const actionKey = Object.keys(LoggerImpl.ACTION).find(k => LoggerImpl.ACTION[k as keyof typeof LoggerImpl.ACTION] === this._action);

    return {
      level: levelKey ?? "UNKNOWN",
      action: actionKey ?? "UNKNOWN",
      errorCount: this._errCnt,
      warningCount: this._warnCnt,
      criticalEvents: [...this._criticalEvents]
    }
  }

  /**
 * Returns the label for a log action value.
 * If no parameter is provided, uses the current logger instance's action.
 * Returns "UNKNOWN" if the value is not found or logger is not initialized.
 */
  public static getActionLabel(action?: typeof LoggerImpl.ACTION[keyof typeof LoggerImpl.ACTION]): string {
    const val = action !== undefined ? action : LoggerImpl._instance?._action
    const UNKNOWN = "UNKNOWN"
    if (val === undefined) return UNKNOWN
    const label = Object.keys(LoggerImpl.ACTION).find(
      key => LoggerImpl.ACTION[key as keyof typeof LoggerImpl.ACTION] === val
    );
    return label ?? UNKNOWN
  }

  /**
  * Returns the label for the given log level.
  * @returns The label for the log level.
  *          If `level` is undefined, returns the label for the current logger instance's level.
  *          If neither is set, returns "UNKNOWN".
  */
  public static getLevelLabel(level?: typeof LoggerImpl.LEVEL[keyof typeof LoggerImpl.LEVEL]): string {
    const val = level !== undefined ? level : LoggerImpl._instance?._level
    const UNKNOWN = "UNKNOWN"
    if (val === undefined) return UNKNOWN
    const label = Object.keys(LoggerImpl.LEVEL).find(
      key => LoggerImpl.LEVEL[key as keyof typeof LoggerImpl.LEVEL] === val
    );
    return label ?? UNKNOWN
  }

  /**
   * Override toString method.
   * @throws ScriptError If the singleton was not instantiated.
   * @override
   */
  public toString(): string {
    const CONTEXT = `${LoggerImpl.name}.toString`
    LoggerImpl.validateInstance(CONTEXT) // Validate the instance
    const NAME = this.constructor.name
    const levelTk = Object.keys(LoggerImpl.LEVEL).find(key =>
      LoggerImpl.LEVEL[key as keyof typeof LoggerImpl.LEVEL] === this._level)
    const actionTk = Object.keys(LoggerImpl.ACTION).find(key =>
      LoggerImpl.ACTION[key as keyof typeof LoggerImpl.ACTION] === this._action)
    const appendersString = Array.isArray(this._appenders)
      ? `[${this._appenders.map(a => a.toString()).join(", ")}]`
      : "[]"
    const scalarInfo = `level: "${levelTk}", action: "${actionTk}", errCnt: ${this._errCnt}, warnCnt: ${this._warnCnt}`
    return `${NAME}: {${scalarInfo}, appenders: ${appendersString}}`
  }

  /**Short version fo the toString() which exludes the appenders details
   * @returns Similar to toString, but showing the list of appenders name only.
   */
  public toShortString(): string {
    const CONTEXT = `${LoggerImpl.name}.toString`
    LoggerImpl.validateInstance(CONTEXT); // Validate the instance
    const NAME = this.constructor.name
    const levelTk = Object.keys(LoggerImpl.LEVEL).find(key =>
      LoggerImpl.LEVEL[key as keyof typeof LoggerImpl.LEVEL] === this._level)
    const actionTk = Object.keys(LoggerImpl.ACTION).find(key =>
      LoggerImpl.ACTION[key as keyof typeof LoggerImpl.ACTION] === this._action)
    const scalarInfo = `level: "${levelTk}", action: "${actionTk}", errCnt: ${this._errCnt}, warnCnt: ${this._warnCnt}`
    const appendersString = Array.isArray(this._appenders)
      ? `[${this._appenders.map(a => a.constructor.name).join(", ")}]`
      : "[]"
    return `${NAME}: {${scalarInfo}, appenders: ${appendersString}}`
  }

  // #TEST-ONLY-START
  /** 
   * Sets the singleton instance to null, useful for running different scenarios.
   * @remarks Mainly intended for testing purposes. The state of the singleton will be lost.
   *          This method only exist in src folder, it wont be deployed in dist folder (production).
   *          It doesn't set the appenders to null, so the appenders are not cleared. 
   * @throws ScriptError If the singleton was not instantiated.
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
    LoggerImpl._instance = null // Force full re-init
  }
  // #TEST-ONLY-END

  /**
 * Routes a log event message (with optional structured extra fields) to all registered appenders.
 * 
 * Behavior:
 * - If the singleton was not instantiated, it is created with default configuration (lazy initialization).
 * - If no appenders are defined, a 'ConsoleAppender' is automatically created and added.
 *   This ensures all logs are delivered to at least one output channel.
 * - The message and extraFields are forwarded to all appenders, which may handle or display them differently (e.g., console, Excel, etc.).
 * - The message is only dispatched if the current log level allows it.
 * - If the event is of type 'ERROR' or 'WARN', it is recorded internally and counted.
 * - If the configured action is 'Logger.ACTION.EXIT', a 'ScriptError' is thrown for errors and warnings.
 * 
 * @param msg - The log message to send.
 * @param type - The log event type (LOG_EVENT).
 * @param extraFields - Optional structured data to attach to the log event (e.g., context info, tags).
 * 
 * @remarks
 * If no appenders are defined, a ConsoleAppender will be created and added automatically.
 * This guarantees that log messages are always delivered to at least one output channel.
 * Extra fields are forwarded to the log event factory and may be included in custom layouts, exports, or external integrations.
 * 
 * @throws ScriptError In case the action defined for the logger is Logger.ACTION.EXIT and the event type
 *         is LOG_EVENT.ERROR or LOG_EVENT.WARN.
 */
  private log(msg: string, type: LOG_EVENT, extraFields?: LogEventExtraFields): void {
    LoggerImpl.initIfNeeded("LoggerImpl.log") // lazy initialization of the singleton with default parameters
    const SEND_EVENTS = (this._level !== LoggerImpl.LEVEL.OFF)
      && (this._level >= type) // Only send events if the level allows it
    if (SEND_EVENTS) {
      if (Utility.isEmptyArray(this._appenders)) {
        this.addAppender(ConsoleAppender.getInstance()) // lazy initialization at least the basic appender
      }
      for (const appender of this._appenders) { // sends to all appenders
        appender.log(msg, type, extraFields) // Pass extraFields through to the appender
      }
      if (type <= LOG_EVENT.WARN) { // Only collects errors or warnings event messages
        // Updating the counter
        if (type === LOG_EVENT.ERROR) ++this._errCnt
        if (type === LOG_EVENT.WARN) ++this._warnCnt
        // Updating the message. Assumes first appender is representative (message for all appenders are the same)
        const appender = this._appenders[0]
        const lastEvent = appender.getLastLogEvent()
        if (!lastEvent) {// internal error
          throw new Error("[LoggerImpl.log] Appender did not return a LogEvent for getLastLogEvent()")
        }
        this._criticalEvents.push(lastEvent)
        if (this._action === LoggerImpl.ACTION.EXIT) {
          const LAST_MSG = AbstractAppender.getLayout().format(lastEvent)
          throw new ScriptError(LAST_MSG)
        }
      }
    }
  }

  /* Enforces instantiation lazily. If the user didn't invoke getInstance(), provides a logger
   * with default configuration. It also sends a trace event indicating the lazy initialization */
  private static initIfNeeded(context?: string): void {
    const PREFIX = context ? `[${context}]: ` : `[LoggerImpl.initIfNeeded]: `
    if (!LoggerImpl._instance) {
      LoggerImpl._instance = LoggerImpl.getInstance()
      const LEVEL_LABEL = `Logger.LEVEL.${LoggerImpl.getLevelLabel()}`
      const ACTION_LABEL = `Logger.ACTION.${LoggerImpl.getActionLabel()}`
      const MSG = `${PREFIX}Logger instantiated via Lazy initialization with default parameters (level='${LEVEL_LABEL}', action='${ACTION_LABEL}')`
      LoggerImpl._instance.trace(MSG)
    }
  }

  // Common safeguard method, where calling initIfNeeded doesn't make sense
  private static validateInstance(context?: string): void {
    if (!LoggerImpl._instance) {
      const PREFIX = context ? `[${context}]: ` : `[${LoggerImpl.name}.validateInstance]: `
      const MSG = `${PREFIX}A singleton instance can't be undefined or null. Please invoke getInstance first.`
      throw new ScriptError(MSG)
    }
  }

  /* Checks level has one of the valid values. It is required, because the way Logger.LEVEL was built,
  i.e. based on LOG_EVENT, so it doesn't check for non-valid values during compilation. That is not the
  case for Logger.ACTION. */
  private static assertValidLevel(level: typeof LoggerImpl.LEVEL[keyof typeof LoggerImpl.LEVEL], context?: string): void {
    if (!Object.values(LoggerImpl.LEVEL).includes(level)) { // level not part of Logger.LEVEL
      const PREFIX = context ? `[${context}]: ` : `[${LoggerImpl.name}.assertValidLevel]: `
      const MSG = `${PREFIX}The input value level='${level}', was not defined in Logger.LEVEL.`
      throw new ScriptError(MSG)
    }
  }

  /** Validates that all appenders are of unique class types, with no null or undefined entries.
   * The uniqueness is based on the constructor of the appender, .i.e. that two different
   * instances of the same appender class are not allowed
  */
  private static assertUniqueAppenderTypes(appenders: (Appender | null | undefined)[], context?: string): void {
    const PREFIX = context ? `[${context}]: ` : `[${LoggerImpl.name}.assertUniqueAppenderTypes]: `
    if (Utility.isEmptyArray(appenders)) {
      throw new ScriptError(`${PREFIX}Invalid input: the input argument 'appenders' must be a non-null array.`)
    }

    const seen = new Set<Function>(); // ensure unique elements only
    for (const appender of appenders) {
      if (!appender) {
        throw new ScriptError(`${PREFIX}Input argument appenders array contains null or undefined entry.`)
      }
      const ctor = appender.constructor
      if (seen.has(ctor)) {
        const name = ctor.name || "UnknownAppender"
        throw new ScriptError(`${PREFIX}Only one appender of type ${name} is allowed.`)
      }
      seen.add(ctor)
    }
  }

  /**
   * Validates that the LEVEL enum values are strictly increasing.
   * This ensures that the log levels are ordered correctly:
   * OFF < ERROR < WARN < INFO < TRACE.
   * @throws ScriptError if the enum values are not strictly increasing.
   * @remarks This is a safeguard to ensure the integrity of the LEVEL enum.
   */
  private static validateLevelEnumIntegrity(): void {
    const levelVals = Object.values(LoggerImpl.LEVEL).filter(v => typeof v === "number") as number[]
    const logEventVals = Object.values(LOG_EVENT).filter(v => typeof v === "number") as number[]
    // 1. Strictly increasing LEVEL values
    for (let i = 1; i < levelVals.length; ++i) {
      if (levelVals[i] <= levelVals[i - 1]) {
        throw new ScriptError(`[LoggerImpl]: LEVEL enum values must be strictly increasing. Found ${levelVals[i - 1]} before ${levelVals[i]}.`)
      }
    }
    // 2. LOG_EVENT values must all be present in LEVEL (except OFF)
    for (const v of logEventVals) {
      if (!levelVals.includes(v)) {
        throw new ScriptError(`[LoggerImpl]: LOG_EVENT value ${v} not present in LEVEL enum.`)
      }
    }
    // 3. LEVEL must have OFF, and all other values must be in LOG_EVENT (excluding OFF)
    if (!("OFF" in LoggerImpl.LEVEL)) {
      throw new ScriptError(`[LoggerImpl]: LEVEL must have OFF value.`)
    }
    for (const v of levelVals) {
      if (v !== LoggerImpl.LEVEL.OFF && !logEventVals.includes(v)) {
        throw new ScriptError(`[LoggerImpl]: LEVEL value ${v} (not OFF) not present in LOG_EVENT.`)
      }
    }
  }

}

// #endregion LoggerImpl


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

// #endregion logger.ts

