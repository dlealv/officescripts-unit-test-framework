/// <reference path="../office-scripts.d.ts" />
// main-wrapper.ts
//
// - This file acts as the entry point for running all tests in a local or CI environment.
// - It sets up a mock ExcelScript environment, imports the test framework, the logger implementation, and the main test runner.
// - The mock workbook is created so scripts can run as if they were in the real Office Scripts runtime.
// - If a global main function is defined (as in Office Scripts), it is called with the mock workbook to execute tests.

import { ExcelScript } from "../mocks/excelscript.mock";
import "../test/unit-test-framework";
import "../src/logger";
import "../test/main";

// Create a mock workbook with two sheets: "Log" and "Sheet1"
const workbook = new ExcelScript.Workbook(["Log", "Sheet1"]);

// Call main with the mock workbook if it is defined
if (typeof main === "function") {
  main(workbook);
}