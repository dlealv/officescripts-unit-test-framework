// main test file for the unit test framework

// #region main.ts

// ----------------------------------------
// Testing unit-test-framework
// ----------------------------------------

// main test file for the unit test framework
function main(workbook: ExcelScript.Workbook
) {

  // Parameters and constants definitions
  // ------------------------------------

  //const VERBOSITY = TestRunner.VERBOSITY.OFF        // uncomment the scenario of your preference
  const VERBOSITY = TestRunner.VERBOSITY.HEADER
  //const VERBOSITY = TestRunner.VERBOSITY.SECTION
  //const VERBOSITY = TestRunner.VERBOSITY.SUBSECTION
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

    // TestRunner tests
    run.title(`${START_TEST} Testing TestRunner Class`, 2)
    run.exec("TestRunner.titleVerbosityOff", () => TestRunnerTest.titleVerbosityOff(), indent)
    run.exec("TestRunner.titleVerbosityOff", () => TestRunnerTest.titlesAndExec, indent)
    run.exec("TestRunner.titleVerbosityOff", () => TestRunnerTest.verbosityProperties, indent)
    run.title(`${END_TEST} Testing TestRunner Class`, 2)

    run.title(`${START_TEST} Testing Assert Class`, 2)
    run.exec("Assert.isTrue", () => AssertTest.isTrue(), indent)
    run.exec("Assert.isFalse", () => AssertTest.isFalse(), indent)
    run.exec("Assert.throws", () => AssertTest.throws(), indent)
    run.exec("Assert.doesNotThrow", () => AssertTest.doesNotThrow(), indent)
    run.exec("Assert.isNull", () => AssertTest.isNull(), indent)
    run.exec("Assert.isNotNull", () => AssertTest.isNotNull(), indent)
    run.exec("Assert.isType", () => AssertTest.isType(), indent)
    run.exec("Assert.equalsPrimitivesAndObjects", () => AssertTest.equalsPrimitivesAndObjects(), indent)
    run.exec("Assert.equalsArrays", () => AssertTest.equalsArrays(), indent)
    run.exec("Assert.instanceOf", () => AssertTest.isInstanceOf(), indent)
    run.exec("Assert.isNotInstanceOf", () => AssertTest.isNotInstanceOf(), indent)
    run.exec("Assert.notEquals", () => AssertTest.notEquals(), indent)
    run.exec("Assert.contains", () => AssertTest.contains(), indent)
    run.exec("Assert.isNotUndefined_and_isDefined", () => AssertTest.isNotUndefined_and_isDefined(), indent)


    // Testing stringify
    run.exec("Test Case AssertSafeStringifyTest.throwsToString", () => AssertSafeStringifyTest.throwsToString(), indent+1)
    run.exec("AssertSafeStringifyTest.circularReference", () => AssertSafeStringifyTest.circularReference(), indent+1)
    run.exec("AssertSafeStringifyTest.symbolValue", () => AssertSafeStringifyTest.symbolValue(), indent+1)
    run.exec("AssertSafeStringifyTest.functionValue", () => AssertSafeStringifyTest.functionValue(), indent+1)
    run.exec("AssertSafeStringifyTest.stringIsQuoted", () => AssertSafeStringifyTest.stringIsQuoted(), indent+1)
    run.exec("AssertSafeStringifyTest.falsyValues", () => AssertSafeStringifyTest.falsyValues(), indent+1)
    run.exec("AssertSafeStringifyTest.null", () => AssertSafeStringifyTest.safeStringify_null(), indent+1)
    run.title(`${END_TEST} Testing Assert Class`, 2)

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
    if (!(e instanceof AssertionError)) { // Unexpected error
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
      console.log(`AssertionError RAISED`)
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
    run.title(success ? `${END_TEST}: OK` : `${END_TEST}: FAIL`, 1)
  }
} // End of main

// Testing Classes
// -----------------

/**Encapsulates the test cases to be executed as static methods of this class. To be
* executed via TestRunner.exec method.
*/
class AssertTest {

  public static isTrue(): void {
    // Positive: Should not throw for truthy values
    Assert.isTrue(true, "isTrue: true should pass")
    Assert.isTrue(1, "isTrue: 1 should pass")
    Assert.isTrue("test", "isTrue: 'test' should pass")
    Assert.isTrue([], "isTrue: [] should pass")
    Assert.isTrue({}, "isTrue: {} should pass")
    Assert.isTrue("non-empty", "isTrue: non-empty string should pass")
    Assert.isTrue(42, "isTrue: 42 should pass")

    // Negative: Should throw for falsy values
    let threw = false
    let errMsg: string

    // null
    threw = false
    try {
      Assert.isTrue(null, "isTrue: null should fail isTrue")
    } catch (e) {
      threw = true
      errMsg = "isTrue: null should fail isTrue: Expected value to be truthy, but got (null)"
      if (!(e instanceof AssertionError)) {
        throw new Error("Did not throw AssertionError for null")
      }
      Assert.isTrue(e.message === errMsg, "Did throw AssertionError but with wrong message for null")
    }
    if (!threw) {
      throw new Error("Assert.isTrue(null) did not throw")
    }

    // undefined
    threw = false
    try {
      Assert.isTrue(undefined, "isTrue: undefined should fail isTrue")
    } catch (e) {
      threw = true
      errMsg = "isTrue: undefined should fail isTrue: Expected value to be truthy, but got (undefined)"
      if (!(e instanceof AssertionError)) {
        throw new Error("Did not throw AssertionError for undefined")
      }
      Assert.isTrue(e.message === errMsg, "Did throw AssertionError but with wrong message for undefined")
    }
    if (!threw) {
      throw new Error("Assert.isTrue(undefined) did not throw")
    }

    // false
    threw = false
    try {
      Assert.isTrue(false, "isTrue: false should fail isTrue")
    } catch (e) {
      threw = true
      errMsg = "isTrue: false should fail isTrue: Expected value to be truthy, but got (false)"
      if (!(e instanceof AssertionError)) {
        throw new Error("Did not throw AssertionError for false")
      }
      Assert.isTrue(e.message === errMsg, "Did throw AssertionError but with wrong message for false")
    }
    if (!threw) {
      throw new Error("Assert.isTrue(false) did not throw")
    }

    // 0
    threw = false
    try {
      Assert.isTrue(0, "isTrue: 0 should fail isTrue")
    } catch (e) {
      threw = true
      errMsg = "isTrue: 0 should fail isTrue: Expected value to be truthy, but got (0)"
      if (!(e instanceof AssertionError)) {
        throw new Error("Did not throw AssertionError for 0")
      }
      Assert.isTrue(e.message === errMsg, "Did throw AssertionError but with wrong message for 0")
    }
    if (!threw) {
      throw new Error("Assert.isTrue(0) did not throw")
    }

    // Empty string
    threw = false
    try {
      Assert.isTrue("", "isTrue: '' should fail isTrue")
    } catch (e) {
      threw = true
      errMsg = "isTrue: '' should fail isTrue: Expected value to be truthy, but got (\"\")"
      if (!(e instanceof AssertionError)) {
        throw new Error("Did not throw AssertionError for empty string")
      }
      Assert.isTrue(e.message === errMsg, "Did throw AssertionError but with wrong message for empty string")
    }
    if (!threw) {
      throw new Error("Assert.isTrue('') did not throw")
    }
  }

  public static isFalse(): void {
    // Positive: should not throw for falsy values
    Assert.isFalse(false, "isFalse: false should pass")
    Assert.isFalse(0, "isFalse: 0 should pass")
    Assert.isFalse("", "isFalse: empty string should pass")
    Assert.isFalse(null, "isFalse: null should pass")
    Assert.isFalse(undefined, "isFalse: undefined should pass")
    Assert.isFalse(NaN, "isFalse: NaN should pass")

    // Negative: should throw for truthy values
    let threw = false
    try {
      Assert.isFalse(true, "isFalse: true should fail")
    } catch (e) {
      threw = true
      if (!(e instanceof AssertionError)) {
        throw new Error("Did not throw AssertionError for true")
      }
    }
    if (!threw) {
      throw new Error("Assert.isFalse(true) did not throw")
    }

    threw = false
    try {
      Assert.isFalse(1, "isFalse: 1 should fail")
    } catch (e) {
      threw = true
      if (!(e instanceof AssertionError)) {
        throw new Error("Did not throw AssertionError for 1")
      }
    }
    if (!threw) {
      throw new Error("Assert.isFalse(1) did not throw")
    }

    threw = false
    try {
      Assert.isFalse("non-empty", "isFalse: 'non-empty' should fail")
    } catch (e) {
      threw = true
      if (!(e instanceof AssertionError)) {
        throw new Error("Did not throw AssertionError for 'non-empty'")
      }
    }
    if (!threw) {
      throw new Error("Assert.isFalse('non-empty') did not throw")
    }

    // Optional: add for [], {}
    threw = false
    try {
      Assert.isFalse([], "isFalse: [] should fail")
    } catch (e) {
      threw = true
      if (!(e instanceof AssertionError)) {
        throw new Error("Did not throw AssertionError for []")
      }
    }
    if (!threw) {
      throw new Error("Assert.isFalse([]) did not throw")
    }

    threw = false
    try {
      Assert.isFalse({}, "isFalse: {} should fail")
    } catch (e) {
      threw = true
      if (!(e instanceof AssertionError)) {
        throw new Error("Did not throw AssertionError for {}")
      }
    }
    if (!threw) {
      throw new Error("Assert.isFalse({}) did not throw")
    }
  }

  public static throws(): void {
    // Positive: should not throw for a function that throws
    Assert.throws(
      () => { throw new Error("Test error") },
      Error,
      "Test error",
      "throws: should catch Error with correct message"
    )

    // Should not throw if the thrown error type matches expected
    Assert.throws(
      () => { throw new AssertionError("Custom assertion failure") },
      AssertionError,
      "Custom assertion failure",
      "throws: should catch AssertionError with correct message"
    )

    // Should not throw if only error type matches and no message is checked
    Assert.throws(
      () => { throw new TypeError("Some type error") },
      TypeError,
      undefined,
      "throws: should catch TypeError"
    )

    // Should not throw if only message matches and no error type is checked
    Assert.throws(
      () => { throw new Error("Just the message") },
      undefined,
      "Just the message",
      "throws: should catch any Error with matching message"
    )

    // Negative: should throw AssertionError if function does NOT throw
    let threw = false
    try {
      Assert.throws(
        () => { let x = 1 + 2 }, // does not throw
        Error,
        undefined,
        "throws: should fail if function does not throw"
      )
    } catch (e) {
      threw = true
      Assert.isTrue(e instanceof AssertionError, "throws: did not throw AssertionError when function did not throw")
    }
    if (!threw) {
      throw new Error("Assert.throws did not throw when function did not throw")
    }

    // Negative: should throw AssertionError if the error type does not match
    threw = false
    try {
      Assert.throws(
        () => { throw new Error("Wrong type") },
        TypeError,
        undefined,
        "throws: should fail if error type does not match"
      )
    } catch (e) {
      threw = true
      Assert.isTrue(e instanceof AssertionError, "throws: did not throw AssertionError when error type did not match")
    }
    if (!threw) {
      throw new Error("Assert.throws did not throw when error type did not match")
    }

    // Negative: should throw AssertionError if the error message does not match
    threw = false
    try {
      Assert.throws(
        () => { throw new Error("Different message") },
        Error,
        "Expected message",
        "throws: should fail if error message does not match"
      )
    } catch (e) {
      threw = true
      Assert.isTrue(e instanceof AssertionError, "throws: did not throw AssertionError when error message did not match")
    }
    if (!threw) {
      throw new Error("Assert.throws did not throw when error message did not match")
    }

    // Negative: should throw AssertionError if thrown value is not an Error
    threw = false
    try {
      Assert.throws(
        () => { throw 123 },
        undefined,
        undefined,
        "throws: should fail if thrown value is not Error"
      )
    } catch (e) {
      threw = true
      Assert.isTrue(e instanceof AssertionError, "throws: did not throw AssertionError when thrown value was not Error")
    }
    if (!threw) {
      throw new Error("Assert.throws did not throw when thrown value was not Error")
    }
  }

  public static doesNotThrow(): void {
    // Positive: should NOT throw for a function that does not throw
    Assert.doesNotThrow(
      () => { let x = 1 + 1 },
      "doesNotThrow: should not throw for simple math"
    )

    Assert.doesNotThrow(
      () => { /* empty function */ },
      "doesNotThrow: should not throw for empty lambda"
    )

    // Negative: should throw AssertionError if the function DOES throw
    let threw = false
    try {
      Assert.doesNotThrow(
        () => { throw new Error("Unexpected error") },
        "doesNotThrow: should fail if function throws"
      )
    } catch (e) {
      threw = true
      Assert.isTrue(e instanceof AssertionError, "doesNotThrow: did not throw AssertionError when function threw")
    }
    if (!threw) {
      throw new Error("Assert.doesNotThrow did not throw when function threw")
    }

    // Negative: thrown value is not an Error (should still catch)
    threw = false
    try {
      Assert.doesNotThrow(
        () => { throw 42 },
        "doesNotThrow: should fail if thrown value is not Error"
      )
    } catch (e) {
      threw = true
      Assert.isTrue(e instanceof AssertionError, "doesNotThrow: did not throw AssertionError for non-Error thrown value")
    }
    if (!threw) {
      throw new Error("Assert.doesNotThrow did not throw when thrown value was not Error")
    }
  }

  public static isNull(): void {
    // Positive: should not throw for null
    Assert.doesNotThrow(
      () => { Assert.isNull(null, "isNull: null should pass") },
      "isNull: should not throw for null"
    )

    // Negative: should throw for non-null values
    Assert.throws(
      () => { Assert.isNull(undefined, "isNull: undefined should fail") },
      AssertionError,
      "isNull: undefined should fail: Expected value to be null, but got (undefined)",
      "isNull: should throw AssertionError for undefined"
    )

    Assert.throws(
      () => { Assert.isNull(0, "isNull: 0 should fail") },
      AssertionError,
      "isNull: 0 should fail: Expected value to be null, but got (0)",
      "isNull: should throw AssertionError for 0"
    )

    Assert.throws(
      () => { Assert.isNull(false, "isNull: false should fail") },
      AssertionError,
      "isNull: false should fail: Expected value to be null, but got (false)",
      "isNull: should throw AssertionError for false"
    )

    Assert.throws(
      () => { Assert.isNull("", "isNull: '' should fail") },
      AssertionError,
      "isNull: '' should fail: Expected value to be null, but got (\"\")",
      "isNull: should throw AssertionError for empty string"
    )

    Assert.throws(
      () => { Assert.isNull("test", "isNull: 'test' should fail") },
      AssertionError,
      "isNull: 'test' should fail: Expected value to be null, but got (\"test\")",
      "isNull: should throw AssertionError for non-empty string"
    )
  }

  public static isNotNull(): void {
    // Positive: should not throw for non-null values
    Assert.doesNotThrow(
      () => { Assert.isNotNull(0, "isNotNull: 0 should pass") },
      "isNotNull: should not throw for 0"
    )

    Assert.doesNotThrow(
      () => { Assert.isNotNull("something", "isNotNull: string should pass") },
      "isNotNull: should not throw for string"
    )

    Assert.doesNotThrow(
      () => { Assert.isNotNull(false, "isNotNull: false should pass") },
      "isNotNull: should not throw for false"
    )

    Assert.doesNotThrow(
      () => { Assert.isNotNull({}, "isNotNull: object should pass") },
      "isNotNull: should not throw for object"
    )

    // Negative: should throw for null
    Assert.throws(
      () => { Assert.isNotNull(null, "isNotNull: null should fail") },
      AssertionError,
      "isNotNull: null should fail: Expected value not to be null, but got (null)",
      "isNotNull: should throw AssertionError for null"
    )
  }

  public static isType(): void {
    // Positive cases: should not throw for correct types
    Assert.doesNotThrow(
      () => { Assert.isType("hello", "string") },
      "assertType: should not throw for string"
    )

    Assert.doesNotThrow(
      () => { Assert.isType(42, "number") },
      "assertType: should not throw for number"
    )

    Assert.doesNotThrow(
      () => { Assert.isType(true, "boolean") },
      "assertType: should not throw for boolean"
    )

    Assert.doesNotThrow(
      () => { Assert.isType({}, "object") },
      "assertType: should not throw for object"
    )

    Assert.doesNotThrow(
      () => { Assert.isType(undefined, "undefined") },
      "assertType: should not throw for undefined"
    )

    Assert.doesNotThrow(
      () => { Assert.isType(Symbol("s"), "symbol") },
      "assertType: should not throw for symbol"
    )

    // Negative cases: should throw for wrong types
    Assert.throws(
      () => { Assert.isType("hello", "number") },
      AssertionError,
      `Expected type 'number', but got 'string': ("hello")`,
      "assertType: should throw for string as number"
    )

    Assert.throws(
      () => { Assert.isType(42, "string") },
      AssertionError,
      `Expected type 'string', but got 'number': (42)`,
      "assertType: should throw for number as string"
    )

    Assert.throws(
      () => { Assert.isType(false, "object") },
      AssertionError,
      `Expected type 'object', but got 'boolean': (false)`,
      "assertType: should throw for boolean as object"
    )

    Assert.throws(
      () => { Assert.isType({}, "number") },
      AssertionError,
      `Expected type 'number', but got 'object': ({})`,
      "assertType: should throw for object as number"
    )

    Assert.throws(
      () => { Assert.isType(undefined, "object") },
      AssertionError,
      `Expected type 'object', but got 'undefined': (undefined)`,
      "assertType: should throw for undefined as object"
    )
  }

  /** Test Assert.isNotType method. */
  public static isNotType(): void {
    // Positive cases (should NOT throw)
    Assert.isNotType("hello", "number", "String is not number")
    Assert.isNotType(42, "string", "Number is not string")
    Assert.isNotType(true, "object", "Boolean is not object")
    Assert.isNotType(undefined, "boolean", "undefined is not boolean")
    Assert.isNotType(Symbol("sym"), "number", "Symbol is not number")
    Assert.isNotType(() => { }, "object", "Function is not object")

    // Negative cases (should throw)
    Assert.throws(
      () => Assert.isNotType("test", "string"),
      AssertionError,
      undefined,
      "Should throw: type matches string"
    )
    Assert.throws(
      () => Assert.isNotType(123, "number"),
      AssertionError,
      undefined,
      "Should throw: type matches number"
    )
    Assert.throws(
      () => Assert.isNotType({}, "object"),
      AssertionError,
      undefined,
      "Should throw: type matches object"
    )
    Assert.throws(
      () => Assert.isNotType(null, "object"),
      AssertionError,
      undefined,
      "null is typeof object in JS"
    )
  }

  // Testing for equality of primitives and objects
  public static equalsPrimitivesAndObjects(): void {
    // Positive: should not throw for equal numbers
    Assert.doesNotThrow(
      () => { Assert.equals(5, 5, "equals: numbers should match") },
      "equals: should not throw for equal numbers"
    )

    // Positive: should not throw for equal strings
    Assert.doesNotThrow(
      () => { Assert.equals("abc", "abc", "equals: strings should match") },
      "equals: should not throw for equal strings"
    )

    // Positive: should not throw for equal booleans
    Assert.doesNotThrow(
      () => { Assert.equals(true, true, "equals: booleans should match") },
      "equals: should not throw for equal booleans"
    )

    // Positive: should not throw for equal objects (deep, via JSON.stringify)
    Assert.doesNotThrow(
      () => { Assert.equals({ a: 1, b: "x" }, { a: 1, b: "x" }, "equals: objects should match") },
      "equals: should not throw for equal objects"
    )

    // Positive: should not throw for null equality
    Assert.doesNotThrow(
      () => { Assert.equals(null, null, "equals: nulls should match") },
      "equals: should not throw for null equality"
    )

    // Positive: should not throw for equal arrays of numbers
    Assert.doesNotThrow(
      () => { Assert.equals([1, 2, 3], [1, 2, 3], "equals: arrays of numbers should match") },
      "equals: should not throw for equal arrays of numbers"
    )

    // Positive: should not throw for equal arrays of objects
    Assert.doesNotThrow(
      () => { Assert.equals([{ x: 1 }, { y: 2 }], [{ x: 1 }, { y: 2 }], "equals: arrays of objects should match") },
      "equals: should not throw for equal arrays of objects"
    )

    // Positive: should not throw for nested arrays/objects
    Assert.doesNotThrow(
      () => { Assert.equals([{ a: [1, 2] }, { b: 3 }], [{ a: [1, 2] }, { b: 3 }], "equals: nested arrays/objects should match") },
      "equals: should not throw for nested arrays/objects equality"
    )

    // Negative: should throw for different numbers
    Assert.throws(
      () => { Assert.equals(5, 6, "equals: numbers should not match") },
      AssertionError,
      "equals: numbers should not match: Assertion failed: actual (5 : number) !== expected (6 : number)",
      "equals: should throw for different numbers"
    )

    // Negative: should throw for different strings
    Assert.throws(
      () => { Assert.equals("abc", "def", "equals: strings should not match") },
      AssertionError,
      "equals: strings should not match: Assertion failed: actual (\"abc\" : string) !== expected (\"def\" : string)",
      "equals: should throw for different strings"
    )

    // Negative: should throw for different booleans
    Assert.throws(
      () => { Assert.equals(true, false, "equals: booleans should not match") },
      AssertionError,
      "equals: booleans should not match: Assertion failed: actual (true : boolean) !== expected (false : boolean)",
      "equals: should throw for different booleans"
    )

    // Negative: should throw for different objects (deep, via JSON.stringify)
    Assert.throws(
      () => { Assert.equals({ a: 1, b: "x" }, { a: 2, b: "x" }, "equals: objects should not match") },
      AssertionError,
      "equals: objects should not match: Assertion failed: actual ({\"a\":1,\"b\":\"x\"}) !== expected ({\"a\":2,\"b\":\"x\"})",
      "equals: should throw for different objects"
    )

    // Negative: should throw for different arrays (length mismatch)
    Assert.throws(
      () => { Assert.equals([1, 2], [1, 2, 3], "equals: arrays length should not match") },
      AssertionError,
      "equals: arrays length should not match: Array length mismatch: actual (2) !== expected (3)",
      "equals: should throw for arrays with different lengths"
    )

    // Negative: should throw for different arrays (element mismatch)
    Assert.throws(
      () => { Assert.equals([1, 2, 4], [1, 2, 3], "equals: arrays element should not match") },
      AssertionError,
      "equals: arrays element should not match: Array value mismatch at index 2: actual (4) !== expected (3)",
      "equals: should throw for arrays with different elements"
    )

    // Negative: should throw for arrays of objects (element mismatch)
    Assert.throws(
      () => { Assert.equals([{ x: 1 }], [{ x: 2 }], "equals: arrays of objects should not match") },
      AssertionError,
      "equals: arrays of objects should not match: Array object value mismatch at index 0: actual ({\"x\":1}) !== expected ({\"x\":2})",
      "equals: should throw for arrays of objects with different field values"
    )

    // Negative: should throw for null vs undefined
    Assert.throws(
      () => { Assert.equals(null, undefined, "equals: null vs undefined should not match") },
      AssertionError,
      "equals: null vs undefined should not match: Assertion failed: actual (null) !== expected (undefined)",
      "equals: should throw for null vs undefined"
    )

    // Negative: should throw for object vs array
    Assert.throws(
      () => { Assert.equals({ 0: 1 }, [1], "equals: object vs array should not match") },
      AssertionError,
      "equals: object vs array should not match: Assertion failed: actual ({\"0\":1}) !== expected ([1])",
      "equals: should throw for object vs array"
    )

    // Negative: should throw for number vs string
    Assert.throws(
      () => { Assert.equals<unknown>(1, "1", "equals: number vs string should not match") },
      AssertionError,
      "equals: number vs string should not match: Assertion failed: actual (1 : number) !== expected (\"1\" : string)",
      "equals: should throw for number vs string"
    )

    // Negative: should throw for nested array/object value mismatch
    Assert.throws(
      () => { Assert.equals([{ a: [1, 2] }, { b: 3 }], [{ a: [1, 2] }, { b: 4 }], "equals: nested arrays/objects should not match") },
      AssertionError,
      "equals: nested arrays/objects should not match: Array object value mismatch at index 1: actual ({\"b\":3}) !== expected ({\"b\":4})",
      "equals: should throw for nested array/object value mismatch"
    )
  }

  // Testing for equal arrays
  public static equalsArrays(): void {
    // Positive: should not throw for equal number arrays
    Assert.doesNotThrow(
      () => { Assert.equals([1, 2, 3], [1, 2, 3], "equals: number arrays should match") },
      "equals: should not throw for equal number arrays"
    )

    // Positive: should not throw for equal string arrays
    Assert.doesNotThrow(
      () => { Assert.equals(["a", "b"], ["a", "b"], "equals: string arrays should match") },
      "equals: should not throw for equal string arrays"
    )

    // Positive: should not throw for arrays of equal objects
    Assert.doesNotThrow(
      () => { Assert.equals([{ x: 1 }, { y: 2 }], [{ x: 1 }, { y: 2 }], "equals: array of objects should match") },
      "equals: should not throw for equal array of objects"
    )

    // Negative: should throw for arrays with different lengths
    Assert.throws(
      () => { Assert.equals([1, 2], [1, 2, 3], "equals: arrays of different length") },
      AssertionError,
      "equals: arrays of different length: Array length mismatch: actual (2) !== expected (3)",
      "equals: should throw for arrays of different length"
    )

    // Negative: should throw for arrays with different values
    Assert.throws(
      () => { Assert.equals([1, 2, 3], [1, 2, 4], "equals: arrays with one different value") },
      AssertionError,
      "equals: arrays with one different value: Array value mismatch at index 2: actual (3) !== expected (4)",
      "equals: should throw for arrays with different values"
    )

    // Negative: should throw for arrays of objects with different values
    Assert.throws(
      () => { Assert.equals([{ x: 1 }], [{ x: 2 }], "equals: arrays of objects with different values") },
      AssertionError,
      "equals: arrays of objects with different values: Array object value mismatch at index 0: actual ({\"x\":1}) !== expected ({\"x\":2})",
      "equals: should throw for arrays of objects with different values"
    )

    // Negative: should throw for arrays with type mismatch
    Assert.throws(
      () => { Assert.equals([1, 2], ["1", "2"], "equals: arrays with type mismatch") },
      AssertionError,
      "equals: arrays with type mismatch: Array type mismatch at index 0: actual (1 : number) !== expected (\"1\" : string)",
      "equals: should throw for arrays with type mismatch"
    )
  }

  // Tests for Assert.isInstanceOf
  public static isInstanceOf(): void {
    class A { }
    class B extends A { }
    class C { }

    // Direct instance
    const a = new A()
    Assert.isInstanceOf(a, A, "a should be instance of A")

    // Inherited instance
    const b = new B()
    Assert.isInstanceOf(b, A, "b should be instance of A (inherited)")
    Assert.isInstanceOf(b, B, "b should be instance of B")

    // Negative: instance of a different class
    const c = new C()
    Assert.throws(
      () => Assert.isInstanceOf(c, A, "c should NOT be instance of A"),
      AssertionError,
      undefined,
      "isInstanceOf: C is not A"
    )

    // Negative: primitive value is not an instance of any class
    Assert.throws(
      () => Assert.isInstanceOf(123, A, "primitive is not instance"),
      AssertionError,
      undefined,
      "isInstanceOf: primitive"
    )

    // Negative: null is not an instance of any class
    Assert.throws(
      () => Assert.isInstanceOf(null, A, "null is not instance"),
      AssertionError,
      undefined,
      "isInstanceOf: null"
    )

    // Negative: undefined is not an instance of any class
    Assert.throws(
      () => Assert.isInstanceOf(undefined, A, "undefined is not instance"),
      AssertionError,
      undefined,
      "isInstanceOf: undefined"
    )

    // Negative: constructor is not a function
    Assert.throws(
      () => Assert.isInstanceOf(a, undefined as unknown as Function, "ctor undefined"),
      AssertionError,
      undefined,
      "isInstanceOf: ctor undefined"
    )
  }

  // Tests for Assert.isNotInstanceOf
  public static isNotInstanceOf(): void {
    class X { }
    class Y extends X { }
    class Z { }

    const x = new X()
    const y = new Y()
    const z = new Z()

    // x should not be instance of Y or Z
    Assert.isNotInstanceOf(x, Y, "x should NOT be instance of Y")
    Assert.isNotInstanceOf(x, Z, "x should NOT be instance of Z")

    // b should not be instance of Z
    Assert.isNotInstanceOf(y, Z, "y should NOT be instance of Z")

    // z should not be instance of X or Y
    Assert.isNotInstanceOf(z, X, "z should NOT be instance of X")
    Assert.isNotInstanceOf(z, Y, "z should NOT be instance of Y")

    // Negative: x is instance of X
    Assert.throws(
      () => Assert.isNotInstanceOf(x, X, "x IS instance of X"),
      AssertionError,
      undefined,
      "isNotInstanceOf: x IS instance of X"
    )

    // Negative: y is instance of X (by inheritance)
    Assert.throws(
      () => Assert.isNotInstanceOf(y, X, "y IS instance of X"),
      AssertionError,
      undefined,
      "isNotInstanceOf: y IS instance of X (by inheritance)"
    )

    // Negative: y is instance of Y
    Assert.throws(
      () => Assert.isNotInstanceOf(y, Y, "y IS instance of Y"),
      AssertionError,
      undefined,
      "isNotInstanceOf: y IS instance of Y"
    )

    // Non-object: primitives, null, undefined are never instances
    Assert.isNotInstanceOf(null, X, "null should NOT be instance of X")
    Assert.isNotInstanceOf(undefined, X, "undefined should NOT be instance of X")
    Assert.isNotInstanceOf(123, X, "primitive should NOT be instance of X")

    // Negative: constructor is not a function
    Assert.throws(
      () => Assert.isNotInstanceOf(x, undefined as unknown as Function, "ctor undefined"),
      AssertionError,
      undefined,
      "isNotInstanceOf: ctor undefined"
    )
  }

    /** Test Assert.isUndefined method. */
  public static isUndefined(): void {
    Assert.isUndefined(undefined, "undefined is undefined")
    Assert.isUndefined(void 0, "void 0 is undefined")
    Assert.throws(
      () => Assert.isUndefined(null),
      AssertionError,
      undefined,
      "null is not undefined"
    )
    Assert.throws(
      () => Assert.isUndefined(0),
      AssertionError,
      undefined,
      "0 is not undefined"
    )
    Assert.throws(
      () => Assert.isUndefined(""),
      AssertionError,
      undefined,
      "empty string is not undefined"
    )
    Assert.throws(
      () => Assert.isUndefined(false),
      AssertionError,
      undefined,
      "false is not undefined"
    )
  }

  /** Test Assert.isNotUndefined and isDefined methods. */
  public static isNotUndefined_and_isDefined(): void {
    Assert.isNotUndefined(0, "0 is not undefined")
    Assert.isNotUndefined(null, "null is not undefined")
    Assert.isNotUndefined(false, "false is not undefined")
    Assert.isDefined(1, "1 is defined")
    Assert.isDefined("", "empty string is defined")
    Assert.isDefined(null, "null is defined")
    Assert.throws(
      () => Assert.isNotUndefined(undefined),
      AssertionError,
      undefined,
      "undefined should throw"
    )
    Assert.throws(
      () => Assert.isDefined(undefined),
      AssertionError,
      undefined,
      "undefined should throw"
    )
  }

  /** Test Assert.notEquals method. */
  public static notEquals(): void {
    Assert.notEquals(1, 2, "1 !== 2")
    Assert.notEquals("foo", "bar", "different strings")
    Assert.notEquals([1, 2], [2, 1], "different arrays")
    Assert.notEquals({ a: 1 }, { a: 2 }, "different objects")
    Assert.throws(
      () => Assert.notEquals(1, 1),
      AssertionError,
      undefined,
      "1 == 1 should throw"
    )
    Assert.throws(
      () => Assert.notEquals("abc", "abc"),
      AssertionError,
      undefined,
      "identical strings"
    )
    Assert.throws(
      () => Assert.notEquals(null, null),
      AssertionError,
      undefined,
      "null == null"
    )
    Assert.throws(
      () => Assert.notEquals([1, 2], [1, 2]),
      AssertionError,
      undefined,
      "equal arrays"
    )
    Assert.throws(
      () => Assert.notEquals({ x: 1 }, { x: 1 }),
      AssertionError,
      undefined,
      "deep equal objects"
    )
  }

  /** Test Assert.contains method for arrays and strings. */
  public static contains(): void {
    // Array contains
    Assert.contains([1, 2, 3], 2, "array contains 2")
    Assert.contains(["a", "b"], "a", "array contains 'a'")
    Assert.throws(
      () => Assert.contains([1, 2, 3], 4),
      AssertionError,
      undefined,
      "array does not contain 4"
    )

    // String contains
    Assert.contains("hello world", "world", "'hello world' contains 'world'")
    Assert.contains("abc", "a", "'abc' contains 'a'")
    Assert.throws(
      () => Assert.contains("abc", "z"),
      AssertionError,
      undefined,
      "'abc' does not contain 'z'"
    )

    // Error: not array or string
    Assert.throws(
      () => Assert.contains(123 as unknown as string, "1"),
      AssertionError,
      undefined,
      "non-array/string container"
    )
    Assert.throws(
      () => Assert.contains({ x: 1 } as unknown as string, "x"),
      AssertionError,
      undefined,
      "object is not valid container"
    )
  }

}

/** 
 * Test coverage for Assert.safeStringify via various assertion failures that
 * surface its behavior. This ensures that error messages are robust for 
 * unusual/hostile values.
 */
class AssertSafeStringifyTest {

  /**Object with toString that throws an error*/
  public static throwsToString(): void {
    // Create an object with a circular reference and a toString that throws
    const obj: { self?: unknown, toString?: () => string } = {}
    obj.self = obj
    obj.toString = function () { throw new Error("Boom!") }

    Assert.throws(
      () => { Assert.equals(obj, {}) },
      AssertionError,
      undefined,
      "safeStringify: should handle object with throwing toString"
    )
  }

  // Test that safeStringify(null) returns 'null' in error messages
  public static safeStringify_null(): void {
    Assert.throws(
      () => { Assert.isNotNull(null) },
      AssertionError,
      'Expected value not to be null, but got (null)',
      "safeStringify: null should stringify to 'null'"
    )
  }

  /**Object with a circular reference*/
  public static circularReference(): void {
    // Create a circular reference without using 'any'
    const a: { self?: unknown } = {}
    a.self = a
    Assert.throws(
      () => { Assert.equals(a, {}) },
      AssertionError,
      undefined,
      "safeStringify: should handle circular reference"
    )
  }


  /**Symbol value*/
  public static symbolValue(): void {
    const sym = Symbol("test")
    Assert.throws(
      () => { Assert.isNull(sym) },
      AssertionError,
      undefined,
      "safeStringify: should handle symbol"
    )
  }

  /**Function value*/
  public static functionValue(): void {
    function f() { }
    Assert.throws(
      () => { Assert.isNull(f) },
      AssertionError,
      undefined,
      "safeStringify: should handle function"
    )
  }

  /**String is quoted in output*/
  public static stringIsQuoted(): void {
    Assert.throws(
      () => { Assert.isNull("abc") },
      AssertionError,
      'Expected value to be null, but got ("abc")',
      "safeStringify: string should be quoted"
    )
  }

  /**Falsy but defined values (0, false, undefined, NaN)*/
  public static falsyValues(): void {
    Assert.throws(
      () => { Assert.isNull(0) },
      AssertionError,
      'Expected value to be null, but got (0)',
      "safeStringify: should stringify 0"
    )
    Assert.throws(
      () => { Assert.isNull(false) },
      AssertionError,
      'Expected value to be null, but got (false)',
      "safeStringify: should stringify false"
    )
    Assert.throws(
      () => { Assert.isNull(undefined) },
      AssertionError,
      'Expected value to be null, but got (undefined)',
      "safeStringify: should stringify undefined"
    )
    Assert.throws(
      () => { Assert.isNull(NaN) },
      AssertionError,
      'Expected value to be null, but got (NaN)',
      "safeStringify: should stringify NaN"
    )
  }
}

class TestRunnerTest {
  /**Test that runnerOff.title does not print (Node: capture output, Office Script: just run)*/
  public static titleVerbosityOff(): void {
    const runnerOff = new TestRunner(TestRunner.VERBOSITY.OFF)
    let captured = ""
    let canOverride = false

    // Try to check if we can override console.log (Node/TS/VSCode only)
    try {
      const origLog = console.log
      console.log = function (msg: string) {
        captured += msg
      }
      canOverride = true
      runnerOff.title("This should not be visible", 1)
      Assert.equals(captured, "", "No output should be printed when verbosity is OFF")
      // Restore
      console.log = origLog
    } catch (e) {
      // Office Script: cannot override console.log, just call the method and not fail
      runnerOff.title("This should not be visible", 1)
    }
  }

  /** Test the title and exec methods of TestRunner, capturing output where possible. */
  public static titlesAndExec(): void {
    // Use a local instance to test verbosity
    const runnerOff = new TestRunner(TestRunner.VERBOSITY.OFF)
    const runnerHeader = new TestRunner(TestRunner.VERBOSITY.HEADER)
    const runnerSection = new TestRunner(TestRunner.VERBOSITY.SECTION)

    let logs: string[] = []
    let canCapture = false
    let originalLog: ((...args: unknown[]) => void) | undefined = undefined

    try {
      // Try to override console.log (Node/VSCode only)
      originalLog = console.log
      console.log = function () {
        // Convert arguments to strings and join for capture
        let out = ""
        for (let i = 0; i < arguments.length; i++) {
          if (i > 0) out += " "
          out += String(arguments[i])
        }
        logs.push(out)
      }
      canCapture = true
    } catch (e) {
      canCapture = false
    }

    // Should NOT print anything
    runnerOff.title("This should not be visible", 1)
    runnerOff.title("This should not be visible either", 2)
    if (canCapture) {
      Assert.equals(logs.length, 0, "runnerOff should not print any titles")
    }

    // Should print only at indent 1
    logs = []
    runnerHeader.title("Header Only", 1)
    runnerHeader.title("Section (should not print)", 2)
    if (canCapture) {
      Assert.equals(logs.length, 1, "runnerHeader should print only one title")
      Assert.isTrue(logs[0].indexOf("Header Only") !== -1, "runnerHeader output contains 'Header Only'")
    }

    // Should print both
    logs = []
    runnerSection.title("Header", 1)
    runnerSection.title("Section", 2)
    if (canCapture) {
      Assert.equals(logs.length, 2, "runnerSection should print two titles")
      Assert.isTrue(logs[0].indexOf("Header") !== -1, "runnerSection output contains 'Header'")
      Assert.isTrue(logs[1].indexOf("Section") !== -1, "runnerSection output contains 'Section'")
    }

    // Restore original log for exec tests (so you can see assertion output if running interactively)
    if (canCapture && originalLog) {
      console.log = originalLog
    }

    // Test exec with a simple passing function
    runnerHeader.exec("Exec Pass", () => {
      Assert.equals(1, 1, "Exec should run this test")
    }, 2)

    // Test exec with a function that fails
    Assert.throws(
      () => runnerHeader.exec("Exec Fail", () => Assert.equals(1, 2, "Should fail")),
      AssertionError,
      undefined,
      "TestRunner.exec should propagate assertion errors"
    )

    // Test exec with a non-function argument
    Assert.throws(
      () => runnerHeader.exec("Not a function", null as unknown as () => void),
      AssertionError,
      undefined,
      "TestRunner.exec should throw if input is not a function"
    )

    // Always restore console.log (for safety)
    if (canCapture && originalLog) {
      console.log = originalLog
    }
  }

  /** Test the getVerbosity and getVerbosityLabel methods of TestRunner. */
  public static verbosityProperties(): void {
    const runner = new TestRunner(TestRunner.VERBOSITY.SECTION)
    Assert.equals(runner.getVerbosity(), TestRunner.VERBOSITY.SECTION, "getVerbosity should return SECTION")
    Assert.equals(runner.getVerbosityLabel(), "SECTION", "getVerbosityLabel should return 'SECTION'")
    Assert.equals(runner.getVerbosityLabel(), "SUBSECTION", "getVerbosityLabel should return 'SUBSECTION'")
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

//#endregion main.ts
