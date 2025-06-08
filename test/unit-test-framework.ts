//****************************************************
// Unit Testing Framework
//****************************************************

// ====================================================
// Lightweight unit testing framework for Office Script
// ====================================================

/**
 * Lightweight, extensible unit testing framework for Office Scripts, inspired by libraries like jUnit.
 * It anable basic assertion and also defines how the test cases are executed.
 * /

/** Utility class for writing unit-test-style assertions.
 * Provides static methods to assert value equality and exception throwing.
 * If an assertion fails, an informative 'Error' is thrown.
 */
class Assert {
  /**
   * Asserts that the provided function throws an error.
   * Optionally checks the error type and message.
   * @param fn - A function that is expected to throw an error.
   *             Must be passed as a function reference, e.g. '() => codeThatThrows()'.
   * @param expectedErrorType - (Optional) Expected constructor of the thrown error (e.g., 'TypeError').
   * @param expectedMessage - (Optional) Exact expected error message.
   * @param message - (Optional) Additional prefix for the error message if the assertion fails.
   * @throws Error - If no error is thrown, or if the thrown error does not match the expected type or message.
   *
   * @example
   * ```ts
   * Assert.throws(() => {
   *   throw new TypeError("Invalid")
   * }, TypeError, "Invalid", "Should throw TypeError")
   * ```
   */
  public static throws(
    fn: () => void,
    expectedErrorType?: Function,
    expectedMessage?: string,
    message: string = ""
  ): asserts fn is () => never {
    const MSG = message ? `${message}: ` : ""

    try {
      fn()
    } catch (e: unknown) {
      if (!(e instanceof Error)) {
        throw new Error(`${MSG} Thrown value is not an Error instance.`)
      }

      if (expectedErrorType && !(e instanceof expectedErrorType)) {
        throw new Error(`${MSG} Expected error type ${expectedErrorType.name}, but got ${e.constructor.name}.`)
      }

      if (expectedMessage && e.message !== expectedMessage) {
        throw new Error(`${MSG} Expected message "${expectedMessage}", but got "${e.message}".`)
      }

      return // ✅ Test passed
    }

    throw new Error(`${MSG} Expected function to throw, but it did not.`)
  }

  /**
   * Asserts that two values are equal by type and value.
   * Supports comparison of primitive types and one-dimensional arrays.
   *
   * If the values differ, a detailed error is thrown.
   * For arrays, mismatches include index, value, and type.
   *
   * @param actual - The actual value.
   * @param expected - The expected value.
   * @param message - (Optional) Prefix message included in the thrown error on failure.
   * @throws Error - If 'actual' and 'expected' are not equal.
   *
   * @example
   * ```ts
   * Assert.equals(2 + 2, 4, "Simple math")
   * Assert.equals(["a", "b"], ["a", "b"], "Array match")
   * Assert.equala([1,2], ["1",2], "Array mismatch") // Expected to fail by type on element 0
   * ```
   */
  public static equals<T>(actual: T, expected: T, message: string = ""): asserts actual is T {
    const MSG = message ? `${message}: ` : ""

    if ((actual == null || expected == null) && actual !== expected) {
      throw new Error(`${MSG}Assertion failed: actual (${actual}) !== expected (${expected})`)
    }

    if (Array.isArray(actual) && Array.isArray(expected)) {
      this.arraysEqual(actual, expected, MSG)
      return
    }

    if (actual !== expected) {
      const actualType = typeof actual
      const expectedType = typeof expected
      throw new Error(`${MSG}Assertion failed: actual (${actual} : ${actualType}) !== expected (${expected} : ${expectedType})`)
    }
  }

/**
 * Asserts that the given value is strictly 'null'.
 * @param value - The value to test.
 * @param message - Optional message to prefix in case of failure.
 * @throws Error if the value is not exactly 'null'.
 */
  public static isNull(value: unknown, message: string = ""): asserts value is null {
    const MSG = message ? `${message}: ` : ""
    if (value !== null) {
      throw new Error(`${MSG}Expected value to be null, but got (${value})`)
    }
  }

  /**
 * Asserts that the given value is strictly not 'null'.
 * @param value - The value to test.
 * @param message - Optional message to prefix in case of failure.
 * @throws Error if the value is not exactly 'null'.
 */
  public static isNotNull(value: unknown, message: string = ""): asserts value is null {
    const MSG = message ? `${message}: ` : ""
    if (value == null) {
      throw new Error(`${MSG}Expected value to be not null, but got (${value})`)
    }
  }

  /**
   * Asserts that two one-dimensional arrays are equal by type and value.
   * Designed for internal use only.
   * @param a - Actual array.
   * @param b - Expected array.
   * @param message - (Optional) Prefix message for errors.
   * @throws Error - If arrays differ in length, type, or value at any index.
   * @private
   */
  private static arraysEqual<T>(a: T[], b: T[], message: string = ""): boolean {
    const MSG = message ? `${message}: ` : ""

    if (a.length !== b.length) {
      throw new Error(`${MSG}Array length mismatch: actual (${a.length}) !== expected (${b.length})`)
    }

    for (let i = 0; i < a.length; i++) {
      const actualValue = a[i]
      const expectedValue = b[i]
      const actualType = typeof actualValue
      const expectedType = typeof expectedValue

      if (actualType !== expectedType) {
        throw new Error(`${MSG}Array type mismatch at index ${i}: actual (${actualValue} : ${actualType}) !== expected (${expectedValue} : ${expectedType})`)
      }

      if (actualValue !== expectedValue) {
        throw new Error(`${MSG}Array value mismatch at index ${i}: actual (${actualValue}) !== expected (${expectedValue})`)
      }
    }
    return true // for consistency; return value is not used
  }
}

/**
 * A utility class for managing and running test cases with controlled console output.
 * TestRunner' supports configurable verbosity levels and allows structured logging
 * with indentation for better test output organization. It is primarily designed for
 * test cases using assertion methods (e.g., 'Assert.equals', 'Assert.throws').
 * Verbosity can be set via the 'TestRunner.VERBOSITY': constant object (enum pattern)
 * - 'OFF' (0): No output
 * - 'HEADER' (1): High-level section headers only
 * - 'SECTION' (2): Full nested titles
 * Verbosity level is incremental, i.e. allows all logging events with indentation level that is
 * lower or equal than TestRunner.VERBOSITY.
 *
 * @example
 * ```ts
 * const runner = new TestRunner()
 * runner.exec("Simple Length Test", () => {
 *   Assert.equals("test".length, 4)
 * })
 * function sum(a: number, b: number): number {return a + b}
 * const a = 1, b = 2;
 * runner.exec("Sum Test", () => {
 *   Assert.equals(sum(a, b), 3) // test passed
 * })
 *
 * // Example output (if verbosity is set to show headers):
 * // ** START: Sum Test **
 * // ** END: Sum Test **
 * ```
 *
 * @remarks Test functions are expected to use `Assert` methods internally. If an assertion fails,
 * the error will be caught and reported with context.
 */
class TestRunner {
  private static readonly START = "START" as const;
  private static readonly END = "END" as const;
  private static readonly HEADER_TK = "*";

  /**Verbosity level */
  public static readonly VERBOSITY = {
    OFF: 0,
    HEADER: 1,
    SECTION: 2,
  } as const;

  // To facilitate the label associated to the verbosity value.
  private static readonly VERBOSITY_LABELS = Object.entries(TestRunner.VERBOSITY).reduce((acc, [key, value]) => {
    acc[value] = key;
    return acc;
  }, {} as Record<number, string>)

  private static readonly DEFAULT_VERBOSITY = TestRunner.VERBOSITY.HEADER
  private readonly _verbosity: typeof TestRunner.VERBOSITY[keyof typeof TestRunner.VERBOSITY]

  /**Constructs a 'TestRunner' with the specified verbosity level.
   * @param verbosity - One of the values from 'TestRunner.VERBOSITY' (default: 'HEADER')
   */
  public constructor(verbosity: typeof TestRunner.VERBOSITY[keyof typeof TestRunner.VERBOSITY] = TestRunner.DEFAULT_VERBOSITY) {
    this._verbosity = verbosity
  }

  /** Returns the current verbosity level. */
  public getVerbosity(): typeof TestRunner.VERBOSITY[keyof typeof TestRunner.VERBOSITY] {
    return this._verbosity
  }

  /** Returns the corresonding string label for the verbosity level. */
  public getVerbosityLabel() {
    return TestRunner.VERBOSITY_LABELS[this._verbosity]
  }

  /**
   * Conditionally prints a title message based on the configured verbosity.
   * The title is prefixed and suffixed with '*' characters for visual structure.
   * The number of '*' will depend on the indentation level, for 2 it shows
   * '**' as prefix and suffix.
   * @param msg - The message to display
   * @param indent - Indentation level (default: `1`). The indentation level is indicated
   *                with the number of suffix '*'.
   */
  public title(msg: string, indent: number = 1): void {
    if (indent <= this._verbosity) {
      const TOKEN = TestRunner.HEADER_TK.repeat(indent);
      console.log(`${TOKEN} ${msg} ${TOKEN}`)
    }
  }

  /** See detailed JSDoc in class documentation */
  public exec(name: string, fn: () => void, indent: number = 2): void {
    this.title(`${TestRunner.START} ${name}`, indent);
    if (typeof fn !== "function") {
      throw new Error("TestRunner.exec() expects a function as input.");
    }
    fn()
    this.title(`${TestRunner.END} ${name}`, indent)
  }
}

// ===========================================================
// End of Lightweight unit testing framework for Office Script
// ===========================================================

// Make Logger and ConsoleAppender available globally for Node/ts-node test environments
if (typeof globalThis !== "undefined") {
  if (typeof TestRunner !== "undefined") {
    // @ts-ignore
    globalThis.TestRunner = TestRunner
  }
  if (typeof Assert !== "undefined") {
    // @ts-ignore
    globalThis.Assert = Assert
  }
}