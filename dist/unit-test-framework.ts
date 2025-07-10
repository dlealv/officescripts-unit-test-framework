
// #region unit-test-framework.ts

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
 * date 2025-06-26 (creation date)
 * version 1.0.0
 */

// #region AssertionError
/**
 * `AssertionError` is a custom error type used to indicate assertion failures in tests or validation utilities.
 * This error is intended to be thrown by assertion methods (such as those in a custom Assert class) when a condition
 * that should always be true is found to be false. Using a specific `AssertionError` type allows for more precise
 * error handling and clearer reporting in test environments, as assertion failures can be easily distinguished from
 * other kinds of runtime errors.
 * @example
 * ```ts
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
// #endregion AssertionError

// #region Assert
/**
 * Utility class for writing unit-test-style assertions.
 * This class provides a set of static methods to perform common assertions
 * such as checking equality, type, and exceptions, etc.
 */
class Assert {

  // #region throws
  /**
   * Asserts that the provided function throws an error.
   * Optionally checks the error type and message.
   * @param fn - A function that is expected to throw an error.
   *             Must be passed as a function reference, e.g. `() => codeThatThrows()`.
   * @param expectedErrorType - (Optional) Expected constructor of the thrown error (e.g., `TypeError`).
   * @param expectedMessage - (Optional) Exact expected error message.
   * @param message - (Optional) Additional prefix for the error message if the assertion fails.
   * @returns {asserts fn is () => never} - Asserts that 'fn' will throw an error if the assertion passes.
   * @throws AssertionError - If no error is thrown, or if the thrown error does not match the expected type or message.
   * @example
   * ```ts
   * Assert.throws(() => {
   *   throw new TypeError("Invalid")
   * }, TypeError, "Invalid", "Should throw TypeError")
   * ```
   * @see {@link Assert.doesNotThrow} for the opposite assertion.
   */
  public static throws(
    fn: () => void,
    expectedErrorType?: Function,
    expectedMessage?: string,
    message: string = ""
  ): asserts fn is () => never {
    const PREFIX = message ? `${message}: ` : ""
    try {
      fn()
    } catch (e: unknown) {
      if (!(e instanceof Error)) {
        throw new AssertionError(`${PREFIX}Thrown value is not an Error instance: (${Assert.safeStringify(e)})`)
      }

      if (expectedErrorType && !(e instanceof expectedErrorType)) {
        throw new AssertionError(`${PREFIX}Expected error type ${expectedErrorType.name}, but got ${e.constructor.name}.`)
      }

      if (expectedMessage && e.message !== expectedMessage) {
        throw new AssertionError(`${PREFIX}Expected message "${expectedMessage}", but got "${e.message}".`)
      }

      return // ✅ Test passed
    }

    throw new AssertionError(`${PREFIX}Expected function to throw, but it did not.`)
  }
  // #endregion throws

  // #region equals
 /**
 * Asserts that two values are equal by type and value.
 * Supports comparison of primitive types, one-dimensional arrays of primitives,
 * and one-dimensional arrays of objects (deep equality via `JSON.stringify`).
 * If the values differ, a detailed error is thrown.
 * For arrays, mismatches include index, value, and type.
 * For arrays of objects, a shallow comparison using `JSON.stringify` is performed.
 * If a value cannot be stringified (e.g., due to circular references), it is treated as "[unprintable value]" in error messages and object equality checks.
 * @param actual - The actual value.
 * @param expected - The expected value.
 * @param message - (Optional) Prefix message included in the thrown error on failure.
 * @returns {asserts actual is T} - Asserts that 'actual' is of type `T` if the assertion passes.
 * @throws AssertionError - If 'actual' and 'expected' are not equal.
 * @example
 * ```ts
 * Assert.equals(2 + 2, 4, "Simple math")
 * Assert.equals(["a", "b"], ["a", "b"], "Array match")
 * Assert.equals([1, "2"], [1, 2], "Array doesn't match") // Fails
 * Assert.equals([{x:1}], [{x:1}], "Object array match") // Passes
 * Assert.equals([{x:1}], [{x:2}], "Object array mismatch") // Fails
 * ```
 * @see {@link Assert.notEquals} for the opposite assertion.
 */
  public static equals<T>(actual: T, expected: T, message: string = ""): asserts actual is T {
    const PREFIX = message ? `${message}: ` : "";

    if ((actual == null || expected == null) && actual !== expected) {
      throw new AssertionError(`${PREFIX}Assertion failed: actual (${Assert.safeStringify(actual)}) !== expected (${Assert.safeStringify(expected)})`);
    }

    if (Array.isArray(actual) && Array.isArray(expected)) {
      this.arraysEqual(actual, expected, message);
      return;
    }

    // Add this block for plain objects
    if (
      typeof actual === "object" &&
      typeof expected === "object" &&
      actual !== null &&
      expected !== null
    ) {
      let actualStr:string, expectedStr:string
      try {
        actualStr = JSON.stringify(actual)
      } catch {
        actualStr = "[unprintable value]"
      }
      try {
        expectedStr = JSON.stringify(expected)
      } catch {
        expectedStr = "[unprintable value]"
      }
      if (actualStr !== expectedStr) {
        throw new AssertionError(
          `${PREFIX}Assertion failed: actual (${Assert.safeStringify(actual)}) !== expected (${Assert.safeStringify(expected)})`
        )
      }
      return
    }

    if (actual !== expected) {
      const actualType = typeof actual;
      const expectedType = typeof expected;
      throw new AssertionError(
        `${PREFIX}Assertion failed: actual (${Assert.safeStringify(actual)} : ${actualType}) !== expected (${Assert.safeStringify(expected)} : ${expectedType})`
      );
    }
  }
  // #endregion equals

  // #region isNull
  /**
   * Asserts that the given value is strictly `null`.
   * Provides a robust stringification of the value for error messages,
   * guarding against unsafe or error-throwing `toString()` implementations.
   * @param value - The value to test.
   * @param message - Optional message to prefix in case of failure.
   * @returns {asserts value is null} - Narrows the type of 'value' to `null` if the assertion passes.
   * @throws AssertionError if the value is not exactly `null`.
   * @example
   * ```ts
   * Assert.isNull(null, "Value should be null")
   * Assert.isNull(undefined, "Value should not be undefined") // Fails
   * Assert.isNull(0, "Zero is not null") // Fails
   * Assert.isNull(null)
   * Assert.isNull(undefined) // Fails
   * Assert.isNull(0) // Fails
   * ```
   * @see {@link Assert.isDefined} for an alias that checks for defined values (not `null` or `undefined`).
   */
  public static isNull(value: unknown, message: string = ""): asserts value is null {
    const PREFIX = message ? `${message}: ` : ""
    if (value !== null) {
      throw new AssertionError(
        `${PREFIX}Expected value to be null, but got (${Assert.safeStringify(value)})`
      )
    }
  }
  // #endregion isNull

  // #region isNotNull
  /**
   * Asserts that the given value is not `null`.
   * Provides a robust stringification of the value for error messages,
   * guarding against unsafe or error-throwing `toString()` implementations.
   * Narrows the type of 'value' to NonNullable<T> if assertion passes.
   * @param value - The value to test.
   * @param message - Optional message to prefix in case of failure.
   * @returns {asserts value is NonNullable<T>} - Narrows the type of 'value' to NonNullable<T> if the assertion passes.
   * @throws AssertionError if the value is `null`.
   * @example
   * ```ts
   * Assert.isNotNull(42, "Value should not be null")
   * Assert.isNotNull(null, "Value should be null") // Fails
   * ```
   * @see {@link Assert.isNull} for the opposite assertion.
   * @see {@link Assert.isDefined} for an alias that checks for defined values (not `null` or `undefined`).
   */
  public static isNotNull<T>(value: T, message: string = ""): asserts value is NonNullable<T> {
    const PREFIX = message ? `${message}: ` : ""
    if (value === null) {
      throw new AssertionError(
        `${PREFIX}Expected value not to be null, but got (${Assert.safeStringify(value)})`
      )
    }
  }
  // #endregion isNotNull

  // #region fail
  /**
   * Fails the test by throwing an error with the provided message.
   * This method is used to explicitly indicate that a test case has failed,
   * regardless of any conditions or assertions.
   * @param message - (Optional) The failure message to display.
   *                  If not provided, a default "Assertion failed" message is used.
   * @returns {never} - This method never returns, it always throws an error.
   * @throws AssertionError - Always throws an `AssertionError` with the provided message.
   * @example
   * ```ts
   * Assert.fail("This test should not pass")
   * ```
   */
  static fail(message?: string) {
    throw new AssertionError(message || "Assertion failed")
  }
  // #endregion fail

  // #region isType
  /**
  * Asserts that a value is of the specified primitive type.
  * @param value - The value to check.
  * @param type - The expected type as a string (`string`, `number`, etc.)
  * @param message - Optional error message.
  * @returns {void} - This method does not return a value.
  * @throws AssertionError - If the type does not match.
  * @example
  * ```ts
  * isType("hello", "string"); // passes
  * isType(42, "number"); // passes
  * isType({}, "string"); // throws
  * isType([], "object", "Expected an object"); // passes
  * isType(null, "object", "Expected an object"); // throws
  * ```
  * @remarks This method checks the type using `typeof` and throws an `AssertionError` if the type does not match.
  *          It is useful for validating input types in functions or methods.
  *         The `type` parameter must be one of the following strings: `string`, `number`, `boolean`, `object`, `function`, `undefined`, `symbol`, or `bigint`.
  *          If the value is `null`, it will be considered an object, which is consistent with JavaScript's behavior.
  * @see {@link Assert.isNotType} for the opposite assertion.
  */
  public static isType(
    value: unknown,
    type: "string" | "number" | "boolean" | "object" | "function" | "undefined" | "symbol" | "bigint",
    message: string = ""
  ): void {
    const PREFIX = message ? `${message}: ` : ""
    if (typeof value !== type) {
      throw new AssertionError(
        `${PREFIX}Expected type '${type}', but got '${typeof value}': (${JSON.stringify(value)})`
      );
    }
  }
  // #endregion isType

  // #region isNotType
/**
 * Asserts that a value is NOT of the specified primitive type.
 * @param value - The value to check.
 * @param type - The unwanted type as a string (`string`, `number`, etc.)
 * @param message - Optional error message.
 * @returns {void} - This method does not return a value.
 * @throws AssertionError - If the type matches.
 * @example
 * ```ts
 * isNotType("hello", "number"); // passes
 * isNotType(42, "string"); // passes
 * isNotType({}, "object"); // throws
 * isNotType(null, "object", "Should not be object"); // throws (null is object in JS)
 * ```
 * @remarks This method checks the type using `typeof` and throws an `AssertionError` if the type matches.
 *          The `type` parameter must be one of the following strings: `string`, `number`, `boolean`, `object`, `function`, `undefined`, `symbol`, or `bigint`.
 * @see {@link Assert.isType} for the positive assertion.
 */
public static isNotType(
  value: unknown,
  type: "string" | "number" | "boolean" | "object" | "function" | "undefined" | "symbol" | "bigint",
  message: string = ""
): void {
  const PREFIX = message ? `${message}: ` : ""
  if (typeof value === type) {
    throw new AssertionError(
      `${PREFIX}Did not expect type '${type}', but got '${typeof value}': (${JSON.stringify(value)})`
    )
  }
}
// #endregion isNotType

  // #region doesNotThrow
  /**
   * Asserts that the provided function does NOT throw an error.
   * If an error is thrown, an `AssertionError` is thrown with the provided message or details of the error.
   * @param fn - A function that is expected to NOT throw.
   *             Must be passed as a function reference, e.g. `() => codeThatShouldNotThrow()`.
   * @param message - (Optional) Prefix for the error message if the assertion fails.
   * @return {void} - This method does not return a value.
   * @throws AssertionError - If the function throws any error.
   * @example
   * ```ts
   * Assert.doesNotThrow(() => {
   *   const x = 1 + 1
   * }, "Should not throw any error")
   * ```
   * @see {@link Assert.throws} for the opposite assertion.
   */
  public static doesNotThrow(fn: () => void, message: string = ""): void {
    const PREFIX = message ? `${message}: ` : ""
    try {
      fn()
    } catch (e) {
      throw new AssertionError(`${PREFIX}Expected function not to throw, but it threw: ${Assert.safeStringify(e)}`)
    }
  }
  // #endregion doesNotThrow

  // #region isTrue
  /**
   * Asserts that the provided value is truthy.
   * Throws AssertionError if the value is not truthy.
   * @param value - The value to test for truthiness.
   * @param message - (Optional) Message to prefix in case of failure.
   * @returns {asserts value} - Narrows the type of 'value' to its original type if the assertion passes.
   * @throws AssertionError - If the value is not truthy.
   * @example
   * ```ts
   * Assert.isTrue(1 < 2, "Math sanity")
   * Assert.isTrue("non-empty string", "String should be truthy")
   * ```
   * @see {@link Assert.isFalse} for the opposite assertion.
   */
  public static isTrue(value: unknown, message: string = ""): asserts value {
    const PREFIX = message ? `${message}: ` : ""
    if (!value) {
      throw new AssertionError(`${PREFIX}Expected value to be truthy, but got (${Assert.safeStringify(value)})`)
    }
  }
  // #endregion isTrue

  // #region isFalse
  /**
   * Asserts that the provided value is falsy.
   * Throws AssertionError if the value is not falsy.
   * @param value - The value to test for falsiness.
   * @param message - (Optional) Message to prefix in case of failure.
   * @returns {void} - This method does not return a value.
   * @throws AssertionError - If the value is not falsy.
   * @example
   * ```ts
   * Assert.isFalse(1 > 2, "Math sanity")
   * Assert.isFalse(null, "Null should be falsy")
   * Assert.isFalse(undefined, "Undefined should be falsy")
   * Assert.isFalse(false, "Boolean false should be falsy")
   * Assert.isFalse(0, "Zero should be falsy")
   * Assert.isFalse("", "Empty string should be falsy")
   * ```
   * @see {@link Assert.isTrue} for the opposite assertion.
   */
  public static isFalse(value: unknown, message: string = ""): void {
    const PREFIX = message ? `${message}: ` : ""
    if (value) {
      throw new AssertionError(`${PREFIX}Expected value to be falsy, but got (${Assert.safeStringify(value)})`)
    }
  }
  // #endregion isFalse

  // #region isUndefined
  /**
   * Asserts that the given value is strictly `undefined`.
   * Throws AssertionError if the value is not exactly `undefined`.
   * @param value - The value to check.
   * @param message - (Optional) Message to prefix in case of failure.
   * @returns {asserts value is undefined} - Narrows the type of 'value' to `undefined` if the assertion passes.
   * @throws AssertionError - If the value is not `undefined`.
   * @example
   * ```ts
   * Assert.isUndefined(void 0)
   * Assert.isUndefined(undefined)
   * Assert.isUndefined(null, "Null is not undefined") // Fails 
   * ```
   * @see {@link Assert.isNotUndefined} for the opposite assertion.
   * @see {@link Assert.isDefined} for an alias that checks for defined values (not `undefined`).
   */
  public static isUndefined(value: unknown, message: string = ""): asserts value is undefined {
    const PREFIX = message ? `${message}: ` : ""
    if (value !== undefined) {
      throw new AssertionError(`${PREFIX}Expected value to be undefined, but got (${Assert.safeStringify(value)})`)
    }
  }
  // #endregion isUndefined

  // #region isNotUndefined
  /**
   * Asserts that the given value is not `undefined`.
   * Narrows the type to exclude undefined.
   * Throws `AssertionError` if the value is `undefined`.
   * @param value - The value to check.
   * @param message - (Optional) Message to prefix in case of failure.
   * @returns {void} - This method does not return a value.
   * @throws AssertionError - If the value is `undefined`.
   * @example
   * ```ts
   * Assert.isNotUndefined(42, "Value should not be undefined")
   * Assert.isNotUndefined(null, "Null is allowed, but not undefined")
   * Assert.isNotUndefined(42)
   * Assert.isNotUndefined(null)
   * ```
   * @see {@link Assert.isUndefined} for the opposite assertion.
   * @see {@link Assert.isDefined} for an alias that checks for defined values (not `undefined`).
   */
  public static isNotUndefined<T>(value: T, message: string = ""): asserts value is Exclude<T, undefined> {
    const PREFIX = message ? `${message}: ` : ""
    if (value === undefined) {
      throw new AssertionError(`${PREFIX}Expected value not to be undefined, but got undefined`)
    }
  }
  // #endregion isNotUndefined

  // #region isDefined
  /**
   * Asserts that the given value is defined (not `undefined`).
   * Alias for `isNotUndefined` method.
   * @param value - The value to check.
   * @param message - (Optional) Message to prefix in case of failure.
   * @returns {void} - This method does not return a value.
   * @throws AssertionError - If the value is `undefined`.
   * @example
   * ```ts
   * Assert.isDefined(42, "Value should be defined")
   * Assert.isDefined(null, "Null is allowed, but not undefined")
   * Assert.isDefined(42)
   * Assert.isDefined(null)
   * ```
   * @see {@link Assert.isNotUndefined} for the opposite assertion.
   * @see {@link Assert.isUndefined} for an alias that checks for undefined values.
   */
  public static isDefined<T>(value: T, message: string = ""): asserts value is Exclude<T, undefined> {
    Assert.isNotUndefined(value, message)
  }
  // #endregion isDefined

  // #region notEquals
  /**
   * Asserts that two values are not equal (deep comparison).
   * For arrays and objects, uses deep comparison (via `JSON.stringify`).
   * Throws `AssertionError` if the values are equal.
   * @param actual - The actual value.
   * @param notExpected - The value that should NOT match.
   * @param message - (Optional) Message to prefix in case of failure.
   * @returns {void} - This method does not return a value.
   * @throws AssertionError - If values are equal.
   * @example
   * ````ts
   * Assert.notEquals(1, 2, "Numbers should not be equal")
   * Assert.notEquals([1, 2], [2, 1], "Arrays should not be equal")
   * Assert.notEquals({ a: 1 }, { a: 2 }, "Objects should not be equal")
   * Assert.notEquals(1, 2)
   * Assert.notEquals([1,2], [2,1])
   * ```
   * @see {@link Assert.equals} for the opposite assertion.
   */
  public static notEquals<T>(actual: T, notExpected: T, message: string = ""): void {
    const PREFIX = message ? `${message}: ` : ""
    try {
      Assert.equals(actual, notExpected, message)
    } catch {
      return // Passed: values are not equal
    }
    throw new AssertionError(`${PREFIX}Values should not be equal: (${Assert.safeStringify(actual)})`)
  }
  // #endregion notEquals

  // #region contains
  /**
   * Asserts that an array or string contains a specified value or substring.
   * For arrays, uses `indexOf` for shallow equality.
   * For strings, uses `indexOf` for substring check.
   * Throws `AssertionError` if the value is not found.
   * @param container - The array or string to search.
   * @param value - The value (or substring) to search for.
   * @param message - (Optional) Message to prefix in case of failure.
   * @return {void} - This method does not return a value.
   * @throws AssertionError - If the value is not found.
   * @example
   * ```ts
   * Assert.contains([1, 2, 3], 2, "Array should contain 2")
   * Assert.contains("hello world", "world", "String should contain 'world'")
   * Assert.contains([1, 2, 3], 4) // Fails
   * Assert.contains("hello world", "test") // Fails
   * Assert.contains([1,2,3], 2)
   * Assert.contains("hello world", "world")
   * ```
   */
  public static contains(container: unknown[] | string, value: unknown, message: string = ""): void {
    const PREFIX = message ? `${message}: ` : ""
    if (typeof container === "string") {
      if (typeof value !== "string" || container.indexOf(value) === -1) {
        throw new AssertionError(`${PREFIX}String does not contain expected substring (${Assert.safeStringify(value)})`)
      }
      return
    }
    if (Array.isArray(container)) {
      if (container.indexOf(value) === -1) {
        throw new AssertionError(`${PREFIX}Array does not contain expected value (${Assert.safeStringify(value)})`)
      }
      return
    }
    throw new AssertionError(`${PREFIX}Contains only works for arrays or strings`)
  }
  // #endregion contains

  // #region isInstanceOf
  /**
   * Asserts that the value is an instance of the specified constructor.
   * Throws `AssertionError` if not.
   * @param value - The value to check.
   * @param ctor - The class/constructor function.
   * @param message - Optional error message prefix.
   * @return {void} - This method does not return a value.
   * @throws AssertionError - If the value is not an instance of the constructor.
   * @example
   * ```ts
   * class MyClass {}
   * const instance = new MyClass()
   * Assert.isInstanceOf(instance, MyClass, "Should be an instance of MyClass")
   * Assert.isInstanceOf(instance, Object) // Passes, since all classes inherit from Object
   * Assert.isInstanceOf(42, MyClass) // Fails
   * ```
   * @see Assert.isNotInstanceOf for the opposite assertion.
   */
  public static isInstanceOf(
    value: unknown,
    ctor: Function,
    message: string = ""
  ): void {
    const PREFIX = message ? `${message}: ` : ""
    if (typeof ctor !== "function") {
      throw new AssertionError(`${PREFIX}Provided constructor is not a function or class.`)
    }
    if (value == null || (typeof value !== "object" && typeof value !== "function")) {
      throw new AssertionError(
        `${PREFIX}Expected instance of ${ctor.name}, but got (${Assert.safeStringify(value)})`
      )
    }
    if (!(value instanceof ctor)) {
      throw new AssertionError(
        `${PREFIX}Expected value to be instance of ${ctor.name}, but got (${Assert.safeStringify(value)})`
      )
    }
  }
  // #endregion isInstanceOf

  // #region isNotInstanceOf
  /**
   * Asserts that the value is NOT an instance of the specified constructor.
   * Throws `AssertionError` if it is.
   * @param value - The value to check.
   * @param ctor - The class/constructor function.
   * @param message - Optional error message prefix.
   * @return {void} - This method does not return a value.
   * @throws AssertionError - If the value is an instance of the constructor.
   * @example
   * ```ts
   * class MyClass {}
   * const instance = new MyClass()
   * Assert.isNotInstanceOf(instance, String, "Should not be an instance of String")
   * Assert.isNotInstanceOf(instance, MyClass) // Fails
   * Assert.isNotInstanceOf(42, MyClass) // Passes, since 42 is not an instance of MyClass
   * ```
   * @see {@link Assert.isInstanceOf} for the opposite assertion.
   */
  public static isNotInstanceOf(
    value: unknown,
    ctor: Function,
    message: string = ""
  ): void {
    const PREFIX = message ? `${message}: ` : ""
    if (typeof ctor !== "function") {
      throw new AssertionError(`${PREFIX}Provided constructor is not a function or class.`)
    }
    if (value != null && (typeof value === "object" || typeof value === "function") && value instanceof ctor) {
      throw new AssertionError(
        `${PREFIX}Expected value NOT to be instance of ${ctor.name}, but got (${Assert.safeStringify(value)})`
      )
    }
  }
  // #endregion isNotInstanceOf

  // #region arraysEqual
  /**
   * Asserts that two one-dimensional arrays are equal by type and value.
   * Supports arrays of primitives and arrays of objects (shallow comparison via JSON.stringify).
   * Designed for internal use only.
   * @param a - Actual array.
   * @param b - Expected array.
   * @param message - (Optional) Prefix message for errors.
   * @returns {boolean} - Returns true if arrays are equal, otherwise throws an error.
   * @throws AssertionError - If arrays differ in length, type, or value at any index.
   * @private
   */
  private static arraysEqual<T>(a: T[], b: T[], message: string = ""): boolean {
    const PREFIX = message ? `${message}: ` : ""

    if (a.length !== b.length) {
      throw new AssertionError(`${PREFIX}Array length mismatch: actual (${a.length}) !== expected (${b.length})`)
    }

    for (let i = 0; i < a.length; i++) {
      const actualValue = a[i]
      const expectedValue = b[i]
      const actualType = typeof actualValue
      const expectedType = typeof expectedValue

      if (actualType !== expectedType) {
        throw new AssertionError(`${PREFIX}Array type mismatch at index ${i}: actual (${Assert.safeStringify(actualValue)} : ${actualType}) !== expected (${Assert.safeStringify(expectedValue)} : ${expectedType})`)
      }

      if (actualType === "object" && expectedType === "object" && actualValue !== null && expectedValue !== null) {
        if (JSON.stringify(actualValue) !== JSON.stringify(expectedValue)) {
          throw new AssertionError(`${PREFIX}Array object value mismatch at index ${i}: actual (${Assert.safeStringify(actualValue)}) !== expected (${Assert.safeStringify(expectedValue)})`)
        }
        continue
      }

      if (actualValue !== expectedValue) {
        throw new AssertionError(`${PREFIX}Array value mismatch at index ${i}: actual (${Assert.safeStringify(actualValue)}) !== expected (${Assert.safeStringify(expectedValue)})`)
      }
    }
    return true // for consistency; return value is not used
  }
  // #endregion arraysEqual

  // #region safeStringify
    /**
   * Returns a safe string representation of any value, handling cases where
   * toString may throw or misbehave. Used internally by assertion methods.
   * Tries `JSON.stringify`, then `value.toString()`, then `Object.prototype.toString.call(value)`.
   * If all fail, returns `[unprintable value]`.
   * @param value - The value to stringify.
   * @returns A string representation of the value, or "[unprintable value]" if not possible.
   * @private
   */
  private static safeStringify(value: unknown): string {
    try {
      if (typeof value === "string") return `"${value}"`
      if (value && typeof value === "object") {
        // Try JSON.stringify
        try {
          return JSON.stringify(value)
        } catch {
          // Try value.toString() if it's a function
          try {
            if (typeof (value as { toString?: unknown }).toString === "function") {
              return (value as { toString: () => string }).toString()
            }
          } catch {}
          // Try Object.prototype.toString.call(value)
          try {
            return Object.prototype.toString.call(value)
          } catch {}
          // All else fails
          return "[unprintable value]"
        }
      }
      return String(value)
    } catch {
      return "[unprintable value]"
    }
  }
  // #endregion safeStringify

}

// #endregion Assert


// #region TestRunner
/**
 * A utility class for managing and running test cases with controlled console output.
 * TestRunner' supports configurable verbosity levels and allows structured logging
 * with indentation for better test output organization. It is primarily designed for
 * test cases using assertion methods (e.g., `Assert.equals`, `Assert.throws`).
 * Verbosity can be set via the `TestRunner.VERBOSITY`: constant object (enum pattern)
 * - `OFF` (0): No output
 * - `HEADER` (1): High-level section headers only
 * - `SECTION` (2): Full nested titles
 * - `SUBSECTION` (3): Detailed test case titles
 * Verbosity level is incremental, i.e. allows all logging events with indentation level that is
 * lower or equal than `TestRunner.VERBOSITY`.
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
  private static readonly START = "START" as const
  private static readonly END = "END" as const
  private static readonly HEADER_TK = "*"

  /**Verbosity level */
  public static readonly VERBOSITY = {
    OFF: 0,
    HEADER: 1,
    SECTION: 2,
    SUBSECTION: 3
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

  /** See detailed JSDoc in class documentation 
   * @param name - The name of the test case.
   * @param fn - The function containing the test logic. It should contain assertions using `Assert` methods.
   * @param indent - Indentation level for the title (default: `2`). The indentation level is indicated
   *                 with the number of suffix '*'.
   * @throws AssertionError - If an assertion fails within the test function.
   * @example
   * ```ts
   * const runner = new TestRunner()
   * runner.exec("My Test", () => {
   *   Assert.equals(1 + 1, 2)
   * })
   * ```
  */
  public exec(name: string, fn: () => void, indent: number = 2): void {
    this.title(`${TestRunner.START} ${name}`, indent);
    if (typeof fn !== "function") {
      throw new AssertionError("TestRunner.exec() expects a function as input.");
    }
    fn()
    this.title(`${TestRunner.END} ${name}`, indent)
  }
}

// #endregion TestRunner

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

//#endregion unit-test-framework.ts

