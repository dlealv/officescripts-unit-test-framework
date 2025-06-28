# Office Scripts Unit Testing Framework

A lightweight, extensible unit testing framework for [Office Scripts](https://learn.microsoft.com/en-us/office/dev/scripts/) & TypeScript, inspired by libraries like JUnit.  
Provides basic assertion capabilities and a structured test runner for easy test authoring, debugging, and reporting—**usable both in Office Scripts and in local Node/TypeScript environments**.

---

## Features

- **Assert Class**: Rich assertion methods for values, arrays (with type and value checking), exceptions, types, containment, and more.
- **TestRunner**: Structured, hierarchical output with configurable verbosity levels (`OFF`, `HEADER`, `SECTION`, `SUBSECTION`).
- **Compatible**: Runs on both Office Scripts and Node/TypeScript (for local or CI testing).
- **Simple**: No dependencies, no decorators, no runtime imports.
- **Extendable**: Add your own assertions or test conventions easily.

---

## Getting Started

### 1. Clone or copy this repo

Place `unit-test-framework.ts` in your project.  
(Optional: Use `main.ts` as a starting point for your test suite.)

### 2. Write Tests

Define a `TestRunner` and create a test class with static methods, e.g.:

```typescript
runner = new TestRunner(TestRunner.VERBOSITY.SECTION)     // Define the test case runneer and verbosity level
runner.title("Start Testing", 1)                          // Sending the title to console indicating the test started
run.exec("Test Case for math", () => TestCase.math(), 1)  // Executing math method from TestCase
runner.title("End Testing", 1)                            // Sending the title to console indicating the test ended

// Class where to organize all test cases
class TestCase {
  public static math(): void {
    Assert.equals(2 + 2, 4, "Addition works")
    Assert.isTrue(5 > 2, "Greater comparison")
    Assert.throws(() => { throw new Error("fail") }, Error, "fail", "Throw check")
  }
}
```
**Note:** `TestCase` class is not requied, just a way to organize all test cases to be executed via `TestRunner` class.

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
  Assert.equals([1, 2, 3], [1, 2, 3], "Arrays are equal"). // Passes: Arrays are equals. Using optional message
  Assert.equals([1, "2"], [1, 2])                          // Fails: type mismatch at index 1
  Assert.equals([{x:1}], [{x:1}])                          // Passes: objects are deeply equal
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

- `OFF` (0): No output.
- `HEADER` (1): Only top-level section headers.
- `SECTION` (2): Section and higher.
- `SUBSECTION` (3): All titles, including subsections.

**How verbosity and indent work:**  
- Each call to `runner.title("Title", indent)` prints the message with `indent` number of `*` as prefix and suffix (e.g., `** title **` for indent=2).
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
function main(workbook: ExcelScript.Workbook) {
  const runner = new TestRunner(TestRunner.VERBOSITY.SECTION)
  let success = false
  try {
    runner.title("Running All Tests", 1)
    runner.exec("Math Test", () => {
      Assert.equals(2 + 3, 5)
      Assert.notEquals(2 * 2, 5)
      Assert.equals([1, 2], [1, 2], "Array equality")
      Assert.throws(() => { throw new TypeError("fail") }, TypeError, "fail", "Throws test")
    }, 2)
    runner.exec("Null/Undefined Test", () => {
      Assert.isNull(null)
      Assert.isNotNull(0)
      Assert.isUndefined(undefined)
      Assert.isNotUndefined("")
      Assert.isDefined(123)
    }, 2)
    runner.exec("Instance Test", () => {
      class Animal {}
      class Dog extends Animal {}
      const d = new Dog()
      Assert.isInstanceOf(d, Dog)
      Assert.isInstanceOf(d, Animal)
      Assert.throws(() => Assert.isInstanceOf({}, Dog), AssertionError, undefined, "Throws if not instance")
      Assert.isNotInstanceOf({}, Dog)
    }, 2)
    runner.exec("Throws/DoesNotThrow Test", () => {
      const thrower = () => { throw new Error("fail") }
      const nonThrower = () => { return 42 }
      Assert.throws(thrower, Error, "fail", "Should throw error")
      Assert.doesNotThrow(nonThrower, "Should not throw")
    }, 2)
    runner.exec("Type Test", () => {
      Assert.assertType("abc", "string")
      Assert.assertType(123, "number")
      Assert.throws(() => Assert.assertType(123, "string"))
    }, 3) // indent=3 (SUBSECTION)
    success = true
  } finally {
    runner.title(success ? "All Tests Passed" : "Test Failure", 1)
  }
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
* All Tests Passed *
```

- Each title uses `*` characters as prefix/suffix, repeated according to the `indent` parameter.
- A title prints only if its `indent` is less than or equal to the runner's verbosity.
- Example above shows only indent 1 and 2 titles, because verbosity is set to `SECTION` (2).

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
- TypeDoc documentation: [TYPEDOC](docs/typedoc/index.html)

## License

[MIT](LICENCE)
