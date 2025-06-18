// ====================================================
// Lightweight unit testing framework for Office Script
// ====================================================

/**
 * Lightweight, extensible unit testing framework for Office Scripts, inspired by libraries like JUnit.
 * Provides basic assertion capabilities and defines the structure for executing test cases.
 * Designed for easy integration and extension within Office Scripts projects.
 *
 * @remarks See the documentation for the Assert and TestRunner classes for assertion details and test execution control.
 * author David Leal
 * date 2025-06-03
 * version 1.2.0
 */

/**
 * AssertionError is a custom error type used to indicate assertion failures in tests or validation utilities.
 * 
 * This error is intended to be thrown by assertion methods (such as those in a custom Assert class) when a condition
 * that should always be true is found to be false. Using a specific AssertionError type allows for more precise
 * error handling and clearer reporting in test environments, as assertion failures can be easily distinguished from
 * other kinds of runtime errors.
 * 
 * Typical usage:
 * 
 * ```typescript
 * if (actual !== expected) {
 *   throw new AssertionError(`Expected ${expected}, but got ${actual}`)
 * }
 * ```
 * 
 * Features:
 * - Inherits from the built-in Error class.
 * - Sets the error name to "AssertionError" for easier identification.
 * - Accepts a message parameter describing the assertion failure.
 * 
 * This class is intentionally simple—no extra methods or properties are added, to keep assertion failures clear and unambiguous.
 */
class AssertionError extends Error {
  constructor(message: string) {
    super(message)
    this.name = "AssertionError"
  }
}

/**
 * Utility class for writing unit-test-style assertions.
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
   * @throws AssertionError - If no error is thrown, or if the thrown error does not match the expected type or message.
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
        throw new AssertionError(`${MSG}Thrown value is not an Error instance: (${Assert.safeStringify(e)})`)
      }

      if (expectedErrorType && !(e instanceof expectedErrorType)) {
        throw new AssertionError(`${MSG}Expected error type ${expectedErrorType.name}, but got ${e.constructor.name}.`)
      }

      if (expectedMessage && e.message !== expectedMessage) {
        throw new AssertionError(`${MSG}Expected message "${expectedMessage}", but got "${e.message}".`)
      }

      return // ✅ Test passed
    }

    throw new AssertionError(`${MSG}Expected function to throw, but it did not.`)
  }

  /**
   * Asserts that two values are equal by type and value.
   * Supports comparison of primitive types, one-dimensional arrays of primitives,
   * and one-dimensional arrays of objects (shallow comparison via JSON.stringify).
   *
   * If the values differ, a detailed error is thrown.
   * For arrays, mismatches include index, value, and type.
   * For arrays of objects, a shallow comparison using JSON.stringify is performed.
   *
   * @param actual - The actual value.
   * @param expected - The expected value.
   * @param message - (Optional) Prefix message included in the thrown error on failure.
   * @throws AssertionError - If 'actual' and 'expected' are not equal.
   *
   * @example
   * ```ts
   * Assert.equals(2 + 2, 4, "Simple math")
   * Assert.equals(["a", "b"], ["a", "b"], "Array match")
   * Assert.equals([{x:1}], [{x:1}], "Object array match") // Passes
   * Assert.equals([{x:1}], [{x:2}], "Object array mismatch") // Fails
   * ```
   */
  public static equals<T>(actual: T, expected: T, message: string = ""): asserts actual is T {
    const MSG = message ? `${message}: ` : ""

    if ((actual == null || expected == null) && actual !== expected) {
      throw new AssertionError(`${MSG}Assertion failed: actual (${Assert.safeStringify(actual)}) !== expected (${Assert.safeStringify(expected)})`)
    }

    if (Array.isArray(actual) && Array.isArray(expected)) {
      this.arraysEqual(actual, expected, MSG)
      return
    }

    if (actual !== expected) {
      const actualType = typeof actual
      const expectedType = typeof expected
      throw new AssertionError(`${MSG}Assertion failed: actual (${Assert.safeStringify(actual)} : ${actualType}) !== expected (${Assert.safeStringify(expected)} : ${expectedType})`)
    }
  }

  /**
   * Asserts that the given value is strictly `null`.
   * Provides a robust stringification of the value for error messages,
   * guarding against unsafe or error-throwing `toString()` implementations.
   * @param value - The value to test.
   * @param message - Optional message to prefix in case of failure.
   * @throws AssertionError if the value is not exactly `null`.
   */
  public static isNull(value: unknown, message: string = ""): asserts value is null {
    const MSG = message ? `${message}: ` : ""
    if (value !== null) {
      throw new AssertionError(
        `${MSG}Expected value to be null, but got (${Assert.safeStringify(value)})`
      )
    }
  }

  /**
   * Asserts that the given value is not `null`.
   * Provides a robust stringification of the value for error messages,
   * guarding against unsafe or error-throwing `toString()` implementations.
   * @param value - The value to test.
   * @param message - Optional message to prefix in case of failure.
   * @throws AssertionError if the value is `null`.
   */
  public static isNotNull<T>(value: T, message: string = ""): asserts value is NonNullable<T> {
    const MSG = message ? `${message}: ` : ""
    if (value === null) {
      throw new AssertionError(
        `${MSG}Expected value not to be null, but got (${Assert.safeStringify(value)})`
      )
    }
  }

  /**
   * Asserts that the provided object is an instance of the specified class or constructor.
   * Throws an error if the assertion fails.
   *
   * @param obj - The object to check.
   * @param cls - The constructor function (class) to check against.
   * @param message - (Optional) Custom error message to display if the assertion fails.
   */
  static instanceOf(obj: any, cls: Function, message?: string) {
    if (!(obj instanceof cls)) {
      throw new AssertionError(
        message || `Expected object to be instance of ${cls.name}, got ${obj?.constructor?.name}`
      )
    }
  }

  /**
   * Asserts that the provided object is NOT an instance of the specified class or constructor.
   * Throws an error if the assertion fails.
   *
   * @param obj - The object to check.
   * @param cls - The constructor function (class) to check against.
   * @param message - (Optional) Custom error message to display if the assertion fails.
   */
  static notInstanceOf(obj: any, cls: Function, message?: string) {
    if (obj instanceof cls) {
      throw new AssertionError(
        message || `Expected object NOT to be instance of ${cls.name}, but got ${obj.constructor.name}`
      )
    }
  }

  /**
   * Fails the test by throwing an error with the provided message.
   *
   * @param message - (Optional) The failure message to display.
   *                  If not provided, a default "Assertion failed" message is used.
   */
  static fail(message?: string) {
    throw new AssertionError(message || "Assertion failed")
  }

  /**
   * Asserts that a value is of the expected primitive type (e.g. "string", "number")
   * or instance of a class/constructor (e.g. Date, Array, custom class).
   * @param value - The value to check.
   * @param typeOrConstructor - The expected type as a string (e.g. "string", "object") or a constructor function.
   * @param message - Optional custom error message.
   */
  static isType(
    value: unknown,
    typeOrConstructor: string | (new (...args: any[]) => any),
    message?: string
  ): void {
    if (typeof typeOrConstructor === "string") {
      if (typeof value !== typeOrConstructor) {
        throw new AssertionError(
          message ||
          `Expected type '${typeOrConstructor}', but got '${typeof value}': (${Assert.safeStringify(value)})`
        )
      }
    } else if (typeof typeOrConstructor === "function") {
      if (!(value instanceof typeOrConstructor)) {
        const ctorName = typeOrConstructor.name || "unknown"
        throw new AssertionError(
          message ||
          `Expected value to be instance of '${ctorName}', but got '${value?.constructor?.name ?? typeof value}': (${Assert.safeStringify(value)})`
        )
      }
    } else {
      throw new AssertionError("Invalid typeOrConstructor argument.")
    }
  }

  /**
   * Asserts that the provided function does NOT throw an error.
   * If an error is thrown, an AssertionError is thrown with the provided message or details of the error.
   *
   * @param fn - A function that is expected to NOT throw.
   *             Must be passed as a function reference, e.g. '() => codeThatShouldNotThrow()'.
   * @param message - (Optional) Prefix for the error message if the assertion fails.
   *
   * @throws AssertionError - If the function throws any error.
   *
   * @example
   * ```ts
   * Assert.doesNotThrow(() => {
   *   const x = 1 + 1
   * }, "Should not throw any error")
   * ```
   */
  public static doesNotThrow(fn: () => void, message: string = ""): void {
    const MSG = message ? `${message}: ` : ""
    try {
      fn()
    } catch (e) {
      throw new AssertionError(`${MSG}Expected function not to throw, but it threw: ${Assert.safeStringify(e)}`)
    }
  }

  /**
   * Asserts that the provided value is truthy.
   * Throws AssertionError if the value is not truthy.
   *
   * @param value - The value to test for truthiness.
   * @param message - (Optional) Message to prefix in case of failure.
   * @throws AssertionError - If the value is not truthy.
   *
   * @example
   * ```ts
   * Assert.isTrue(1 < 2, "Math sanity")
   * Assert.isTrue("non-empty string", "String should be truthy")
   * ```
   */
  public static isTrue(value: unknown, message: string = ""): asserts value {
    const MSG = message ? `${message}: ` : ""
    if (!value) {
      throw new AssertionError(`${MSG}Expected value to be truthy, but got (${Assert.safeStringify(value)})`)
    }
  }

  /**
   * Asserts that the provided value is falsy.
   * Throws AssertionError if the value is not falsy.
   *
   * @param value - The value to test for falsiness.
   * @param message - (Optional) Message to prefix in case of failure.
   * @throws AssertionError - If the value is not falsy.
   *
   * @example
   * Assert.isFalse(0, "Zero should be falsy")
   * Assert.isFalse("", "Empty string should be falsy")
   */
  public static isFalse(value: unknown, message: string = ""): void {
    const MSG = message ? `${message}: ` : ""
    if (value) {
      throw new AssertionError(`${MSG}Expected value to be falsy, but got (${Assert.safeStringify(value)})`)
    }
  }

  /**
   * Asserts that two one-dimensional arrays are equal by type and value.
   * Supports arrays of primitives and arrays of objects (shallow comparison via JSON.stringify).
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
      throw new AssertionError(`${MSG}Array length mismatch: actual (${a.length}) !== expected (${b.length})`)
    }

    for (let i = 0; i < a.length; i++) {
      const actualValue = a[i]
      const expectedValue = b[i]
      const actualType = typeof actualValue
      const expectedType = typeof expectedValue

      if (actualType !== expectedType) {
        throw new AssertionError(`${MSG}Array type mismatch at index ${i}: actual (${Assert.safeStringify(actualValue)} : ${actualType}) !== expected (${Assert.safeStringify(expectedValue)} : ${expectedType})`)
      }

      if (actualType === "object" && expectedType === "object" && actualValue !== null && expectedValue !== null) {
        if (JSON.stringify(actualValue) !== JSON.stringify(expectedValue)) {
          throw new AssertionError(`${MSG}Array object value mismatch at index ${i}: actual (${Assert.safeStringify(actualValue)}) !== expected (${Assert.safeStringify(expectedValue)})`)
        }
        continue
      }

      if (actualValue !== expectedValue) {
        throw new AssertionError(`${MSG}Array value mismatch at index ${i}: actual (${Assert.safeStringify(actualValue)}) !== expected (${Assert.safeStringify(expectedValue)})`)
      }
    }
    return true // for consistency; return value is not used
  }

  /**
   * Returns a safe string representation of any value, handling cases where
   * toString may throw or misbehave. Used internally by assertion methods.
   * @param value - The value to stringify.
   * @returns A string representation of the value, or a fallback if not possible.
   */
  private static safeStringify(value: unknown): string {
    try {
      if (typeof value === "string") return `"${value}"`
      if (value && typeof value === "object") {
        try {
          return JSON.stringify(value)
        } catch {
          return value.toString?.() ?? Object.prototype.toString.call(value)
        }
      }
      return String(value)
    } catch {
      return "[unprintable value]"
    }
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
      throw new AssertionError("TestRunner.exec() expects a function as input.");
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
  if (typeof AssertionError !== "undefined") {
    // @ts-ignore
    globalThis.AssertionError = AssertionError
  }
}