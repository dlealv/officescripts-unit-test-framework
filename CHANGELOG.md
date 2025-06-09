# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [Unreleased]

---

## [1.2.0] – 2025-06-09

### Changed
- **Documentation:** Updated README and developer guide to clarify that all logger functionality—including `Logger`, `ConsoleAppender`, and `ExcelAppender`—is now implemented in a single file: `src/logger.ts`.
- **Usage Guidance:** Examples and instructions now direct users to copy *only* `src/logger.ts` into their Office Scripts project; all references to separate appender files have been removed. The README now also explains that the `dist` folder contains production-ready files with `clearInstance` methods removed, and that those methods are only present in source files for testing or development—not for production use.
- **Testing Instructions:** Clearly documented that `clearInstance` methods for logger and appenders are available only in source (not production) code, and are intended for test scenarios.
- **TypeScript Configuration:** Updated `tsconfig.json` (`strictNullChecks: false`) to better match Office Scripts runtime behavior, ensuring that local TypeScript execution emulates Office Scripts' permissiveness with `null` and `undefined`.

---

## [1.0.0] – 2025-06-08

### Added
- Initial release of the Office Scripts Logging Framework.
- TypeScript-based logger for Office Scripts with mock/test harness.
- GitHub Actions CI workflow.
- Developer workflow and project documentation.

---
