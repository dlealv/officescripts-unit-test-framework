// Augment global scope for test/mock detection and runtime ExcelScript environment flag.
//
// - This file declares a global variable for use in detecting the execution environment in code and tests.
// - DO NOT declare `var ExcelScript` here, as that would cause a duplicate identifier error with the ExcelScript namespace.
// - In your test setup (not in a .d.ts), you can assign to globalThis.ExcelScript as needed.
// - Use ExcelScriptIsMock to distinguish between Office Scripts and test/runtime environments.

export {}; // Required for proper global augmentation

declare global {
  /**
   * True if running in a mocked/test environment. Used for environment detection.
   * Assign in your test setup: globalThis.RunSyncTest = true to force synchronous execution, i.e. with 
   * no delay between script execution and test assertions.
   */
  var RunSyncTest: boolean | undefined;
  
}