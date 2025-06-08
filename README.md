# Office Scripts Logging Framework – User Guide

A lightweight, extensible logging framework for [Office Scripts](https://learn.microsoft.com/en-us/office/dev/scripts/) (ExcelScript), inspired by Log4j.  
Add robust, structured logs to your Excel automations with control over log levels, appenders, and error handling.

---

## Features

- **Multiple Log Levels:** `ERROR`, `WARN`, `INFO`, `TRACE`, plus an `OFF` mode.
- **Configurable Error Handling:** Continue or terminate script on warnings/errors.
- **Pluggable Appenders:** Output logs to console or Excel cells, or build your own.
- **Singleton Design:** Easy, static usage – only one logger instance per script.
- **TypeScript/Office Scripts Compatible:** Works in both Office Scripts and Node test environments.

---

## Quick Start

### 1. Add the Logger to Your Script

Copy `src/logger.ts` into your Office Scripts project.

### 2. Initialize the Logger

```typescript
// Import or copy the Logger, ConsoleAppender, and ExcelAppender classes

// Set log level and action (optional: see below for defaults)
const logger = Logger.getInstance(Logger.LEVEL.INFO, Logger.ACTION.CONTINUE);

// Add at least one appender (where logs are output):
Logger.addAppender(ConsoleAppender.getInstance()); // Output to console

// OPTIONAL: Output logs to a cell in Excel
// const cell = workbook.getWorksheet("Log").getRange("A1");
// Logger.addAppender(ExcelAppender.getInstance(cell));
```

---

## Usage Examples

### Basic Logging

```typescript
Logger.info("Script started");
Logger.warn("This might be a problem");
Logger.error("A fatal error occurred");
Logger.trace("Step-by-step details for debugging");
```

### Example: Log to Excel Cell

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Setup logger to send logs to cell B1
  const cell = workbook.getActiveWorksheet().getRange("B1");
  Logger.clearInstance(); // (optional: only needed if re-running in same session)
  Logger.getInstance(Logger.LEVEL.INFO, Logger.ACTION.CONTINUE);
  Logger.addAppender(ExcelAppender.getInstance(cell));

  Logger.info("Log written to Excel!");
  Logger.warn("This warning appears in cell B1.");
}
```

---

## Configuration

### Log Levels

Set the **minimum severity** of messages to be logged:

- `Logger.LEVEL.OFF`: No logs
- `Logger.LEVEL.ERROR`: Only errors
- `Logger.LEVEL.WARN`: Errors and warnings
- `Logger.LEVEL.INFO`: Errors, warnings, and info (default)
- `Logger.LEVEL.TRACE`: All messages (most verbose)

### Error Handling Action

Choose how your script responds to errors/warnings:

- `Logger.ACTION.CONTINUE`: Log the event, continue script execution (default)
- `Logger.ACTION.EXIT`: Log the event and throw a `ScriptError`, terminating the script

### Appenders

Where logs go:

- `ConsoleAppender`: Output to the Office Scripts console
  ```typescript
  Logger.addAppender(ConsoleAppender.getInstance());
  ```
- `ExcelAppender`: Output to a specified Excel cell, with color coding
  ```typescript
  Logger.addAppender(ExcelAppender.getInstance(cellRange));
  ```
  - You can only have one appender of each type.

---

## Advanced Usage

### Manage Appenders

- **Add:** `Logger.addAppender(appender)`
- **Remove:** `Logger.removeAppender(appender)`
- **Replace all:** `Logger.setAppenders([appender1, appender2])`

### Inspect Logger State

- **Get all error/warning messages:** `logger.getMessages()`
- **Get error/warning counts:** `logger.getErrCnt()`, `logger.getWarnCnt()`
- **Clear state (not appenders):** `logger.clear()`
- **Export state:** `logger.exportState()`

### Reset Logger (for testing or to change configuration)

```typescript
Logger.clearInstance();
```

---

## API Reference

### Main Methods

| Method             | Description                                           |
|--------------------|------------------------------------------------------|
| `Logger.error()`   | Log error (always logged if level ≥ ERROR)           |
| `Logger.warn()`    | Log warning (if level ≥ WARN)                        |
| `Logger.info()`    | Log info (if level ≥ INFO)                           |
| `Logger.trace()`   | Log trace/debug details (if level ≥ TRACE)           |

### Static Properties

- `Logger.LEVEL`: Log levels (`OFF`, `ERROR`, `WARN`, `INFO`, `TRACE`)
- `Logger.ACTION`: Error-handling actions (`CONTINUE`, `EXIT`)

---

## Frequently Asked Questions

- **Q:** What happens if I don’t add an appender?  
  **A:** The logger will default to `ConsoleAppender` (logs go to console).

- **Q:** Can I log to both console and Excel?  
  **A:** Yes, add both appenders.

- **Q:** How do I change log level or action after initialization?  
  **A:** Use `Logger.clearInstance()` and then call `getInstance()` with new options.

---

## Example: Full Script

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Reset logger and set up
  Logger.clearInstance();
  Logger.getInstance(Logger.LEVEL.TRACE, Logger.ACTION.CONTINUE);

  // Add appenders
  Logger.addAppender(ConsoleAppender.getInstance());
  const logCell = workbook.getActiveWorksheet().getRange("C2");
  Logger.addAppender(ExcelAppender.getInstance(logCell));

  // Logging
  Logger.info("Script started.");
  Logger.trace("This is a trace message.");
  Logger.warn("This is a warning.");
  Logger.error("This is an error!"); // If ACTION.EXIT, this throws and aborts the script
}
```

---

## Errors & Troubleshooting

- If you see a `ScriptError`, check if `Logger.ACTION.EXIT` is set.
- Each appender can only be added once; duplicates throw.
- Always call `Logger.getInstance()` before adding appenders or logging.

---

## License

MIT

---

*For developer setup, testing, or CI details, see [docs/DEVELOPER_GUIDE.md](docs/DEVELOPER_GUIDE.md) if available.*
