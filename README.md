# Office Scripts Logging Framework – User Guide

A lightweight, extensible logging framework for [Office Scripts](https://learn.microsoft.com/en-us/office/dev/scripts/) (ExcelScript), inspired by Log4j.  
Add robust, structured logs to your Excel automations with control over log levels, appenders, and error handling.

---

## Features

- **Multiple Log Levels:** `ERROR`, `WARN`, `INFO`, `TRACE`, plus an `OFF` mode.
- **Configurable Error Handling:** Continue or terminate script on warnings/errors.
- **Pluggable Appenders:** Output logs to console or Excel cells, or build your own.
- **Singleton & Lazy Initialization:** The logger and appenders are created only when first needed.
- **TypeScript/Office Scripts Compatible:** Works in both Office Scripts and Node test environments.

---

## ⚡ Lazy Initialization (How It Works)

- **Logger Singleton:**  
  - You do **not** have to manually create the Logger instance with `Logger.getInstance()` before logging.  
  - If you call any logging method directly (such as `Logger.info("...")`) before explicit initialization, the Logger instance will be **automatically created** with **default settings** (`WARN` level, `EXIT` action).

- **Default ConsoleAppender:**  
  - If you log a message and **no appender** has been added, a `ConsoleAppender` will be **automatically created and added**.
  - This ensures logs are never lost, even if you forget to add an appender.

- **Appenders:**  
  - Appenders are singletons and use lazy initialization, especially `ConsoleAppender` and `ExcelAppender`.  
  - For `ExcelAppender`, the first call requires the cell range; subsequent calls return the existing instance and ignore new arguments.

**Result:**  
You can start using the Logger immediately, but for best control, explicitly initialize it and add appenders as shown below.

---

## Quick Start

### 1. Add the Logger to Your Script

Copy `src/logger.ts` into your Office Scripts project.

### 2. Initialize the Logger (Recommended)

```typescript
Logger.getInstance(Logger.LEVEL.INFO, Logger.ACTION.CONTINUE);
Logger.addAppender(ConsoleAppender.getInstance());
```

> **Tip:** If you skip this step and just call `Logger.info("...")`, the logger will be initialized with default settings and a console appender.

---

## Usage Examples

### Basic Logging (with Lazy Initialization)

```typescript
// This works even if you haven't called getInstance() or added an appender.
// Logger will be auto-initialized at WARN level, EXIT action, with a ConsoleAppender.
Logger.info("Script started");
Logger.warn("This might be a problem");
Logger
