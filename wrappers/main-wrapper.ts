
// main-wrapper.ts
//
// - This file acts as the entry point for running all tests in a local or CI environment.
// - It sets up a mock ExcelScript environment, imports the test framework, the logger implementation, and the main test runner.
// - The mock workbook is created so scripts can run as if they were in the real Office Scripts runtime.
// - If a global main function is defined (as in Office Scripts), it is called with the mock workbook to execute tests.

// Import the Office Scripts mock
import { ExcelScript } from "../mocks/excelscript.mock";

// Import the test framework to register global TestRunner, Assert, etc.
import "../test/unit-test-framework";

// Import the logger (optional, if your main test file expects it)
import "../src/logger";

// Import the actual test suite (should define main or global tests)
import "../test/main";

// Create a mock workbook with two sheets: "Log" and "Sheet1"
const workbook = new ExcelScript.Workbook(["Log", "Sheet1"]);

// Call main with the mock workbook if it is defined globally
// (main may be defined in globalThis, depending on your main.ts)
if (typeof (globalThis as any).main === "function") {
  (globalThis as any).main(workbook);
} else if (typeof main === "function") {
  // Fallback if main is in the current scope (older setup)
  (main as any)(workbook);
} else {
  console.log("No global main function found to execute tests.");
}