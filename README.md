# Office Scripts Logging Framework

A lightweight, extensible logging framework designed for [Office Scripts](https://learn.microsoft.com/en-us/office/dev/scripts/overview/excel).  
This library enables structured logging in Excel online automations, with TypeScript-first design and compatibility for both Office Scripts and local Node.js testing environments.

---

## Features

- **Structured Logging:** Easily add trace, debug, info, warning, and error logs to your Office Scripts.
- **TypeScript Support:** Written in TypeScript for safety and IntelliSense.
- **Office Scripts Compatibility:** Works seamlessly in Excel online automations.
- **Pluggable Outputs:** Designed for easy extension (e.g., log to worksheet, console, etc.).
- **Testable Locally:** Includes mocks for Office Scripts APIs for local/unit testing in Node.js.
- **Minimal Dependencies:** Lightweight and simple to integrate.

---

## Getting Started

### 1. Installation

Clone the repository and install dependencies:

```sh
git clone <your-repo-url>
cd <your-repo-folder>
npm ci
```

### 2. Usage in Office Scripts

Copy the `src/` files or your built JavaScript into the Office Scripts editor in Excel online.

**Basic Example:**

```typescript
// Import or copy Logger from src/
const logger = new Logger("MyScript");
logger.info("Script started.");

// ...your automation logic...

logger.success("Script completed successfully.");
```

See [docs/office-scripts.md](docs/office-scripts.md) for more on integrating with Office Scripts.

---

## API Overview

| Method              | Description                                  |
|---------------------|----------------------------------------------|
| `logger.trace(msg)` | Trace-level log                              |
| `logger.debug(msg)` | Debug-level log                              |
| `logger.info(msg)`  | Info-level log                               |
| `logger.warn(msg)`  | Warning-level log                            |
| `logger.error(msg)` | Error-level log                              |
| `logger.success(msg)` | Success indicator (if implemented)         |

See the code in `src/Logger.ts` for further details and customization options.

---

## Advanced Usage & Extensibility

- **Custom Log Handlers:**  
  Implement your own output targets (e.g., log to worksheet, send to API) by extending the logger or configuring handlers.

- **Configuration:**  
  Set log levels, formats, or targets as needed.

- **TypeScript Typings:**  
  The library is fully typed and compatible with the Office Scripts type system.

See [docs/usage-examples.md](docs/usage-examples.md) and [docs/office-scripts.md](docs/office-scripts.md) for more.

---

## Testing

This project supports full local unit testing using mocks for the Office Scripts API.

### Run All Tests

```sh
npm run build
npm test
```

- The main test entry point is `wrappers/mainWrapper.ts`.
- Mocks for ExcelScript are in `mocks/excelscript.mock.ts`.
- Tests are defined in `test/main.ts` and related files.

See [docs/testing.md](docs/testing.md) for full details on the testing setup.

---

## Project Structure

```
src/                  # Logging framework source code
test/                 # Unit tests (entry: test/main.ts)
wrappers/mainWrapper.ts # Test runner using ExcelScript mocks
mocks/excelscript.mock.ts # Mock implementation for testing
office-scripts.d.ts   # Type definitions for Office Scripts
.github/workflows/    # CI configuration
docs/                 # Supporting documentation
```

---

## Contributing

Contributions are welcome!

- Please see [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.
- For major changes, open an issue to discuss your proposal.

---

## CI/CD

Automated testing and builds run on each push and pull request to `main` via GitHub Actions.  
The workflow checks that TypeScript builds and all tests pass before allowing merging.  
See [docs/ci.md](docs/ci.md) for details.

---

## Further Documentation

- [Developer Guide](docs/DEVELOPER_GUIDE.md)
- [Testing Guide](docs/testing.md)
- [CI/CD Pipeline](docs/ci.md)
- [Office Scripts Integration](docs/office-scripts.md)
- [Usage Examples](docs/usage-examples.md)
- [Changelog](CHANGELOG.md)

---

## License

[MIT](LICENSE)

---

## Support

For questions or issues, please open an issue in this repository.
