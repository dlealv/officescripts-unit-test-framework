# Office Scripts Logger User Guide

A lightweight, extensible logging utility for [Office Scripts](https://learn.microsoft.com/office/dev/scripts/).  
Easily add trace, debug, info, warning, and error logs to your Excel scripts.

---

## Features

- Multiple log levels: trace, debug, info, warn, error, success
- Lightweight and easy to integrate with Office Scripts
- Simple API for adding logs anywhere in your script
- Output logs to the Excel workbook or custom targets

---

## Getting Started

### 1. Installing / Adding to Your Script

**Option A: Copy Source Code**

Copy `Logger.ts` (or the main logger file) from `src/` into the Office Scripts editor in Excel online.

**Option B: Use the Built JavaScript**

If you use a bundler/transpiler, paste the output code into the Office Scripts editor.

---

### 2. Basic Usage Example

```typescript
// If you copied Logger.ts, instantiate the logger
const logger = new Logger("MyOfficeScript");

function main(workbook: ExcelScript.Workbook) {
  logger.info("Script started.");

  // Your script logic
  try {
    // ... do work
    logger.debug("Doing some work...");
    // Simulate action
    logger.success("Work completed successfully!");
  } catch (e) {
    logger.error(`An error occurred: ${e}`);
  }
}
```

---

### 3. Log Levels

- `logger.trace(message)`
- `logger.debug(message)`
- `logger.info(message)`
- `logger.warn(message)`
- `logger.error(message)`
- `logger.success(message)`

Each method will add a log entry at the specified level.

---

### 4. Outputting Logs

**Default:**  
- Logs are stored in memory (or as designed in your Logger).
- Extend the Logger to output to a worksheet, named range, or elsewhere as needed.

**Example: Output logs to a worksheet**

```typescript
// After your script runs, write logs to a worksheet
function main(workbook: ExcelScript.Workbook) {
  const logger = new Logger("MyOfficeScript");

  logger.info("Started");

  // ... your code

  // At the end, output logs
  const sheet = workbook.addWorksheet("Logs");
  sheet.getRange("A1").setValues([["Level", "Message"]]);
  let row = 2;
  for (const entry of logger.getLogs()) {
    sheet.getCell(row - 1, 0).setValue(entry.level);
    sheet.getCell(row - 1, 1).setValue(entry.message);
    row++;
  }
}
```
*(Adjust the API if your Logger has a different log retrieval/output method.)*

---

### 5. Customization

- **Log format:** You can extend or modify the Logger to change timestamp format, add author/script name, etc.
- **Custom outputs:** Implement methods to send logs to URLs, custom sheets, etc.

---

### 6. Office Scripts Compatibility Notes

- Only use APIs available in the [Office Scripts documentation](https://learn.microsoft.com/office/dev/scripts/).
- Do not use Node.js or browser-specific APIs.

---

## Frequently Asked Questions

**Q: Can I use this logger in VBA/macros or desktop Excel?**  
A: No, it is designed for Office Scripts in Excel Online.

**Q: Does it work for Google Sheets App Scripts?**  
A: No, but you could adapt the code for that platform.

**Q: How do I test my script locally?**  
A: See [docs/DEVELOPER_GUIDE.md](docs/DEVELOPER_GUIDE.md) for developer setup and testing info.

---

## License

MIT

---

*For developer setup, contributing, or CI details, see [docs/DEVELOPER_GUIDE.md](docs/DEVELOPER_GUIDE.md).*
