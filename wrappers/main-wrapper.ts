/// <reference path="../office-scripts.d.ts" />
import { ExcelScript } from "../mocks/excelscript.mock";
import "../test/unit-test-framework";
import "../src/logger";
import "../src/main";

// Create a mock workbook
const workbook = new ExcelScript.Workbook(["Log", "Sheet1"]);

// Call main with the mock workbook
if (typeof main === "function") {
  main(workbook);
}