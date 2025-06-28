// File: wrappers/main-wrapper.ts
// - This file acts as the entry point for running all tests in a local or CI environment.
// - If a global main function is defined (as in Office Scripts), it is called with the mock workbook to execute tests.
// - Compatible with both Office Scripts and local Node.js/TypeScript environments.
// - The ExcelScript mock is required for testing without Excel Online.
//
// Usage:
// - Run with ts-node or as part of your CI scripts.
// - Ensures that Office Scripts code is testable outside of Excel Online.

import { ExcelScript } from "../mocks/excelscript.mock"
import "../src/unit-test-framework"
import "../test/main"

const workbook = new ExcelScript.Workbook()

if (typeof (globalThis as any).main === "function") {
  (globalThis as any).main(workbook)
} else {
  console.log("No global main function found to execute tests.")
}