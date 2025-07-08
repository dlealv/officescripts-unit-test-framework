# Office Scripts Unit Testing Framework

A lightweight, extensible unit testing framework for [Office Scripts](https://learn.microsoft.com/en-us/office/dev/scripts/) & TypeScript, inspired by libraries like JUnit.  
Provides basic assertion capabilities and a structured test runner for easy test authoring, debugging, and reporting—**usable both in Office Scripts and in local Node/TypeScript environments**.

---

## Features

- **Assert Class**: Rich assertion methods for values, arrays (with type and value checking), exceptions, types, containment, and more.
- **TestRunner**: Structured, hierarchical output with configurable verbosity levels (`OFF`, `HEADER`, `SECTION`, `SUBSECTION`).
- **Compatible**: Runs on both Office Scripts and Node/TypeScript (for local or CI testing).
- **Simple**: No dependencies, no decorators, no runtime imports.
- **Extensible**: Add your own assertions or test conventions easily.

---

## Getting Started

### 1. Clone or copy this repo

Place `unit-test-framework.ts` in your project.  
(Optional: Use `test/main.ts` as a starting point for your test suite.)

### 2. Write Tests

Define a `TestRunner` and create a test class with static methods, e.g.:

```typescript
const runner = new TestRunner(TestRunner.VERBOSITY.SECTION)     // Define the test case runner and verbosity level
runner.title("Start Testing", 1)                                // Output title indicating the test started
runner.exec("Test Case for math", () => TestCase.math(), 2)     // Execute math method from TestCase with section indentation level
runner.title("End Testing", 1)                                  // Output title indicating the test ended

// Class to organize all test cases
class TestCase {
  public static math(): void {
    Assert.equals(2 + 2, 4, "Addition works")
    Assert.isTrue(5 > 2, "Greater comparison")
    Assert.throws(() => { throw new Error("fail") }, Error, "fail", "Throw check")
  }
}
```
**Note:** The `TestCase` class is not required, just a way to organize all test cases to be executed via the `TestRunner` class.

### 3. Run Tests

In Office Scripts, call `main(workbook)` (see `test/main.ts`).

In Node/TypeScript, run a wrapper (see `main-wrapper.ts`) that invokes `main`.

---

## API Reference

### Assert Class

#### Value Equality & Arrays

```typescript
Assert.equals(actual, expected, "optional message")
```
- **Supports primitives, arrays, and objects.**
- For arrays, each element is checked for both type and value. For objects/arrays of objects, a deep check (using JSON.stringify) is performed.
- Example:
  ```typescript
  Assert.equals([1, 2, 3], [1, 2, 3], "Arrays are equal") // Passes
  Assert.equals([1, "2"], [1, 2])                         // Fails: type mismatch at index 1
  Assert.equals([{x:1}], [{x:1}])                         // Passes: objects are deeply equal
  ```

#### Inequality

```typescript
Assert.notEquals(actual, notExpected, "optional message")
```

#### Instance Checks

```typescript
Assert.isInstanceOf(obj, ClassConstructor, "optional message")
Assert.isNotInstanceOf(obj, ClassConstructor, "optional message")
```

#### Type Checks

```typescript
Assert.isType(value, "string" | "number" | "boolean" | "object" | "function" | "undefined" | "symbol" | "bigint", "optional message")
Assert.isNotType(value, "string" | "number" | "boolean" | "object" | "function" | "undefined" | "symbol" | "bigint", "optional message")
```
- Example:
  ```typescript
  Assert.isType("hello", "string", "Should be string")
  Assert.isType(42, "number")
  Assert.isType({}, "object")
  Assert.isNotType("hello", "number", "String is not number")
  Assert.isNotType(42, "string", "Number is not string")
  ```

#### Null/Undefined Checks

```typescript
Assert.isNull(value, "optional message")
Assert.isNotNull(value, "optional message")
Assert.isUndefined(value, "optional message")
Assert.isNotUndefined(value, "optional message")
Assert.isDefined(value, "optional message") // alias for isNotUndefined
```

#### Truthy/Falsy

```typescript
Assert.isTrue(expression, "optional message")
Assert.isFalse(expression, "optional message")
```

#### Containment

```typescript
Assert.contains(arrayOrString, value, "optional message")
```
- Example:
  ```typescript
  Assert.contains([1, 2, 3], 2, "Array contains 2")
  Assert.contains("hello world", "world", "Substring found")
  ```

#### Exception Assertions

To test that code throws (or does not throw) as expected, always pass a function reference using `() => ...`.  
If you pass a direct function call (e.g., `Assert.throws(myFunction())`), the code will execute before it reaches the assertion and the assertion won't work as intended.

**Example:**

Suppose you have the following simple class:

```typescript
class Divider {
  static divide(a: number, b: number): number {
    if (b === 0) throw new Error("Cannot divide by zero")
    return a / b
  }
}
```

You can test that `Divider.divide` throws for zero denominator, and does not throw otherwise:

```typescript
// Correct: Pass a function reference (using an arrow function)
Assert.throws(
  () => Divider.divide(10, 0),
  Error,
  "Cannot divide by zero",
  "Should throw when dividing by zero"
)

// Also correct: test that a valid division does NOT throw
Assert.doesNotThrow(
  () => Divider.divide(10, 2),
  "Should not throw for valid division"
)
```

**Note:**  
`Assert.throws` requires **the throwing code to be passed as a function reference** (using `() => ...` or `function() { ... }`).  
This allows the assertion method to execute your function and catch any exceptions inside its own logic.

#### Fail Manually

```typescript
Assert.fail("This should not happen")
```

---

### TestRunner Class

#### Creating a Test Runner

```typescript
const runner = new TestRunner(TestRunner.VERBOSITY.SECTION) // or HEADER, OFF, SUBSECTION
```

#### Verbosity Levels

- `OFF` (`0`): No output.
- `HEADER` (`1`): Only top-level section headers.
- `SECTION` (`2`): Section and higher.
- `SUBSECTION` (`3`): All titles, including subsections.

**How verbosity and indent work:**  
- Each call to `runner.title("Title", indent)` prints the message with `indent` number of `*` as prefix and suffix (e.g., `** title **` for `indent=2`).
- A title is only printed if its `indent` is **less than or equal to** the current verbosity.
- This lets you control granularity of test output: higher verbosity shows more detail.

#### Running Tests

```typescript
runner.exec("My Test Name", () => {
  Assert.equals(1 + 1, 2)
}, 2) // The '2' is the indent level for this test (prints if verbosity >= 2)
```

#### Structured Output

```typescript
runner.title("Title the testing", 1) // * Title the testing *
runner.title("Section", 2)           // ** Section **
runner.title("Detail", 3)            // *** Detail ***
```

#### Getting Verbosity

```typescript
runner.getVerbosity()      // returns numeric level
runner.getVerbosityLabel() // returns "HEADER", etc
```

---

## Example: Full Test Suite

```typescript
// main test file for the unit test framework

function main(workbook: ExcelScript.Workbook) {
  const runner = new TestRunner(TestRunner.VERBOSITY.SECTION)
  let success = false
  try {
    runner.title("Running All Tests", 1)
    runner.exec("Math Test", () => TestCase.math(), 2)
    runner.exec("Null/Undefined Test", () => TestCase.nullUndefined(), 2)
    runner.exec("Instance Test", () => TestCase.instance(), 2)
    runner.exec("Throws/DoesNotThrow Test", () => TestCase.throwsDoesNotThrow(), 2)
    runner.exec("Type Test", () => TestCase.type(), 2)
    success = true
  } finally {
    runner.title(success ? "All Tests Passed" : "Test Failure", 1)
  }
}

// Class to organize all test cases as static methods
class TestCase {
  public static math() {
    Assert.equals(2 + 3, 5, "Addition works")
    Assert.notEquals(2 * 2, 5, "Multiplication does not equal 5")
    Assert.equals([1, 2], [1, 2], "Array equality")
  }

  public static nullUndefined() {
    Assert.isNull(null, "Should be null")
    Assert.isNotNull(0, "Zero is not null")
    Assert.isUndefined(undefined, "Should be undefined")
    Assert.isNotUndefined("", "Empty string is defined")
    Assert.isDefined(123, "Number is defined")
  }

  public static instance() {
    class Animal {}
    class Dog extends Animal {}
    const d = new Dog()
    Assert.isInstanceOf(d, Dog, "Dog instance of Dog")
    Assert.isInstanceOf(d, Animal, "Dog instance of Animal")
    Assert.throws(() => Assert.isInstanceOf({}, Dog), AssertionError, undefined, "Throws if not instance")
    Assert.isNotInstanceOf({}, Dog, "Plain object is not instance of Dog")
  }

  public static throwsDoesNotThrow() {
    // --- All throws cases ---
    // 1. Throws an Error with specific message
    Assert.throws(() => { throw new Error("fail") }, Error, "fail", "Should throw Error")

    // 2. Throws a TypeError
    Assert.throws(() => { throw new TypeError("bad type") }, TypeError, "bad type", "Should throw TypeError")

    // 3. Throws any error (not checking error type or message)
    Assert.throws(() => { throw "custom error string" }, undefined, undefined, "Should throw any error (string)")

    // 4. Throws AssertionError when an assertion fails inside
    Assert.throws(() => Assert.isTrue(false, "Forced fail"), AssertionError, undefined, "Should throw AssertionError when assertion fails")

    // 5. Using a function variable that throws
    const failFunc = () => { throw new RangeError("range fail") }
    Assert.throws(failFunc, RangeError, "range fail", "Should throw RangeError")

    // --- All doesNotThrow cases ---
    // 1. Does not throw (simple value)
    Assert.doesNotThrow(() => 42, "Should not throw on returning 42")

    // 2. Does not throw (returns undefined)
    Assert.doesNotThrow(() => undefined, "Should not throw on returning undefined")

    // 3. Does not throw (assertion that passes)
    Assert.doesNotThrow(() => Assert.isTrue(true, "Should pass"), "Should not throw if assertion passes")

    // 4. Using a function variable that does not throw
    const safeFunc = () => "hello"
    Assert.doesNotThrow(safeFunc, "Should not throw with safeFunc")
  }

  public static type() {
    Assert.isType("abc", "string", "abc is string")
    Assert.isType(123, "number", "123 is number")
    Assert.throws(() => Assert.isType(123, "string"), undefined, undefined, "Throws if type mismatch")
    Assert.isNotType("hello", "number", "String is not number")
    Assert.isNotType(42, "string", "Number is not string")
  }
}

// Make main available globally for Node/ts-node test environments
if (typeof globalThis !== "undefined" && typeof main !== "undefined") {
  // @ts-ignore
  globalThis.main = main
}
```

---

## Output Example (Verbosity: SECTION)

```
* Running All Tests *
** START Math Test **
** END Math Test **
** START Null/Undefined Test **
** END Null/Undefined Test **
** START Instance Test **
** END Instance Test **
** START Throws/DoesNotThrow Test **
** END Throws/DoesNotThrow Test **
** START Type Test **
** END Type Test **
* All Tests Passed *
```

- Each title uses `*` characters as prefix/suffix, repeated according to the `indent` parameter.
- A title prints only if its `indent` is less than or equal to the runner's verbosity.
- Example above shows only indent `1` and `2` titles, because verbosity is set to `SECTION` (`2`).

If verbosity level is `HEADER` the output will be:
```
* Running All Tests *
* All Tests Passed *
```

---

## Development & Customization

- Add your own assertion methods to the `Assert` class.
- Organize tests as you wish—group by topic, file, or feature.
- Works directly in the Office Scripts editor (Excel Online), as well as in VSCode/Node with your own mocks.

---

## Additional Information

- TypeDoc documentation: [TYPEDOC](https://dlealv.github.io/officescripts-unit-test-framework/typedoc/)

## License

[MIT](LICENCE)
