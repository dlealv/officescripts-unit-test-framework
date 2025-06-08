# Office Scripts Logging Framework – User Guide

A lightweight, extensible logging framework for [Office Scripts](https://learn.microsoft.com/en-us/office/dev/scripts/) (ExcelScript), inspired by frameworks like Log4j.  
Add robust, structured logs to your Excel automations with control over log levels, appenders, and error handling.

---

## Features

- **Multiple Log Levels:** `ERROR`, `WARN`, `INFO`, `TRACE`, plus an `OFF` mode.
- **Configurable Error Handling:** Continue or terminate the script on warnings/errors.
- **Pluggable Appenders:** Output logs to the console or Excel cells, or build your own.
- **Singleton & Lazy Initialization:** The logger and appenders are created only when first needed.
- **TypeScript/Office Scripts Compatible:** Works in both Office Scripts and Node test environments.

---

## Lazy Initialization: What You Need to Know

- **Logger Singleton:**  
  - You do **not** have to manually create the Logger instance before using it  
  - If you call any logging method (e.g., `Logger.info("...")`) before calling `getInstance()`, the Logger will be **automatically created** with default settings (`WARN` level, `EXIT` action)
- **Default ConsoleAppender:**  
  - If you log a message and **no appender** has been added, a `ConsoleAppender` will be **automatically created and added**. This ensures logs are never lost, even if you forget to add an appender

**Summary:**  
You can start logging immediately, but for best results (and explicit configuration), initialize the Logger and add your desired appenders as shown below

---

## Getting Started

### 1. Add the Logger to Your Script

Copy `src/logger.ts` into your Office Scripts project

### 2. Initialize the Logger (Recommended)

```typescript
// Set verbosity level up to INFO events, and continue on error/warning
Logger.getInstance(Logger.LEVEL.INFO, Logger.ACTION.CONTINUE)
Logger.addAppender(ConsoleAppender.getInstance()) // Add console appender
```
> If you skip this step and just call `Logger.info("...")`, the logger will be created with default settings and a console appender:
> * `Logger` will be initialized with verbosity level up to warnings, and in case of error/warning, execution stops by throwing a `ScriptError`
> * The default appender used will be the `ConsoleAppender`, which doesn't require any configuration input parameters

---

## Usage Examples

### Basic Logging

```typescript
Logger.info("Script started")
Logger.warn("This might be a problem")
Logger.error("A fatal error occurred")
Logger.trace("Step-by-step details for debugging")
```
> Even if you haven’t explicitly initialized the Logger or added an appender, logging will still work (see Lazy Initialization above)

### Logging to Excel Cell

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Set up logger to send logs to cell B1
  const cell = workbook.getActiveWorksheet().getRange("B1")
  Logger.clearInstance() // (optional, if rerunning this script multiple times)
  Logger.getInstance(Logger.LEVEL.INFO, Logger.ACTION.CONTINUE)
  Logger.addAppender(ExcelAppender.getInstance(cell))

  Logger.info("Log written to Excel!")
  Logger.warn("This warning appears in cell B1.")
}
```

---

## Configuration

### Log Levels

Set the **minimum severity** of messages to be logged:

- `Logger.LEVEL.OFF`: No logs
- `Logger.LEVEL.ERROR`: Only errors
- `Logger.LEVEL.WARN`: Errors and warnings (default)
- `Logger.LEVEL.INFO`: Errors, warnings, and info
- `Logger.LEVEL.TRACE`: All messages (most verbose)

### Error Handling Action

- `Logger.ACTION.CONTINUE`: Log the event, continue script execution
- `Logger.ACTION.EXIT`: Log the event and throw a `ScriptError`, terminating the script (default)
  > The configuration above only applies for log events sent to the appenders. If the level is `Logger.LEVEL.OFF`, no log events will be sent to any appender

### Appenders

- `ConsoleAppender`: Output to the Office Scripts console  
  `Logger.addAppender(ConsoleAppender.getInstance())`
- `ExcelAppender`: Output to a specified Excel cell, with color coding  
  `Logger.addAppender(ExcelAppender.getInstance(cellRange))`

---

## Advanced Usage

### Manage Appenders

- Add: `Logger.addAppender(appender)`
- Remove: `Logger.removeAppender(appender)`
- Replace all: `Logger.setAppenders([appender1, appender2])`
- **Only one of each appender type is allowed; duplicates will throw an error**

### Inspect Logger State

- Get an array of all error/warning messages sent to the appenders: `logger.getMessages()`
- Get error/warning counts: `logger.getErrCnt()`, `logger.getWarnCnt()`
- Clear state (messages, counters, but not appenders): `logger.clear()`
- Export state: `logger.exportState()`

### Reset Logger

- Use `Logger.clearInstance()` to reset the singleton and allow new configuration (useful in test loops or if your script reruns in the same session)

---

## API Reference

### Main Methods

| Method                          | Description                                                            |
|----------------------------------|------------------------------------------------------------------------|
| `Logger.error(message: string)`  | Logs error event with a message, if `level >= LEVEL.ERROR`             |
| `Logger.warn(message: string)`   | Logs warning with a message, if `level >= LEVEL.WARN`                  |
| `Logger.info(message: string)`   | Logs info with a message, if `level >= LEVEL.INFO`                     |
| `Logger.trace(message: string)`  | Logs trace/debug details with a message, if `level >= LEVEL.TRACE`     |

### Static Properties

- `Logger.LEVEL`: Log levels (`OFF`, `ERROR`, `WARN`, `INFO`, `TRACE`)
- `Logger.ACTION`: Error-handling actions (`CONTINUE`, `EXIT`)

---

## Complete Example

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Reset logger and set up
  Logger.clearInstance()
  // Set verbosity up to TRACE and continue on error/warning
  Logger.getInstance(Logger.LEVEL.TRACE, Logger.ACTION.CONTINUE)

  // Add appenders
  Logger.addAppender(ConsoleAppender.getInstance())
  const logCell = workbook.getActiveWorksheet().getRange("C2")
  Logger.addAppender(ExcelAppender.getInstance(logCell))

  // Logging
  Logger.info("Script started.")
  Logger.trace("This is a trace message.")
  Logger.warn("This is a warning.")
  Logger.error("This is an error!") // If ACTION.EXIT, this throws and aborts the script
}
```

---

## Troubleshooting & FAQ

- **What if I call Logger methods before getInstance()?**  
  Lazy initialization means logging always works, with default config and console output

- **What happens if I don’t add an appender?**  
  Logger auto-adds a `ConsoleAppender`

- **Can I log to both console and Excel?**  
  Yes, add both appenders

- **How do I change log level or action after initialization?**  
  Use `Logger.clearInstance()` and then call `getInstance()` with new options

- **Why do I get a `ScriptError`?**  
  If `Logger.ACTION.EXIT` is set and `Logger.LEVEL != LEVEL.OFF`, errors/warnings throw and abort the script

- **Why can I only add one of each appender type?**  
  To avoid duplicate logs on the same channel; each appender represents a unique output

- **Why can't I send a different message to different appenders?**  
  By design, all channels (appenders) receive the same log event message for consistency

---

## License

See [LICENSE](LICENSE) for details

---

*For developer setup, testing, or CI details, see [docs/DEVELOPER_GUIDE.md](docs/DEVELOPER_GUIDE.md) if available*
