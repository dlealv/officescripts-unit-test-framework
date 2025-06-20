# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [Unreleased]

---

## [2.0.0] – 2025-06-19

### Added
- **AssertionError**: Introduced a new error class for clearer assertion failure reporting in the unit test framework (`unit-test-framework.ts`).
- **Assert enhancements**: Added new convenience methods to the `Assert` class, making test writing more robust and expressive (`unit-test-framework.ts`).
- **New interfaces and types**: Introduced `Logger`, `Layout`, and `LogEvent` interfaces, as well as the `LogEventFactory` type, to enhance extensibility and type safety (`logger.ts`).
- **New classes**: Added `LoggerImpl`, `Utility`, `LayoutImpl`, `LogEventImpl` (`logger.ts`), and `AssertionError` (`unit-test-framework.ts`) to provide a more modular, extensible, and testable architecture.
- **Layout-based output customization**: Appenders now support customizable log message layouts via the `Layout` interface, which exposes a `format` method. The `LayoutImpl` class allows further customization by accepting a user-defined `formatter` function through its constructor, giving users complete control over log message formatting.
- **`VSCode Debuggins.md`**: Detail documentation on how to debug the library while running tests.

### Changed
- **Logger refactored**: The `Logger` is now an interface, implemented by the new `LoggerImpl` class. Common helper and validation methods have been moved to the `Utility` class for broader reuse.
- **Appender output control**: All appenders now utilize shared formatting logic via the `Layout` interface (managed within `AbstractAppender`). The actual output of log messages is managed by implementing the abstract `sendEvent` method in each appender subclass.
- **ScriptError improvement**: Improved the `ScriptError.raiseIfNeeded` helper method to support custom error handling scenarios.
- **Standardized the output of `toString()` methods** for classes based on best practices. This includes consistent formatting, clear use of public property names, and structured output to improve debugging and testing.  
  Reference: [Best Practices for toString() in JavaScript/TypeScript](https://stackoverflow.com/questions/65358186/best-practices-for-tostring-in-javascript-typescript)
- **Extended the `LogEvent` interface to support a generic type parameter** for custom extra fields, allowing flexible extension of log event metadata. Adjusted the interfaces and the rest of the classes to allow log event with extra parameters.
- **Change license from GNU to MIT**: Updated the `LICENSE` file.


### Breaking
- **Logger API**: The previous `Logger` class has been replaced by a `Logger` interface and a `LoggerImpl` implementation. All usages must be updated to use the new API.
- **Appender APIs**: Output formatting is now manag ed through the `Layout` interface and its `format` method. Custom layouts may require new construction or configuration patterns. Message output must be implemented via `sendEvent` in subclasses of `AbstractAppender`.
- Renamed `Logger.clear()` to `Logger.reset()` for clarity. The method now clearly indicates it resets only the error/warning counters and critical event messages, but does not affect singleton instances, appenders, layout, or log event factory. This change improves code readability and avoids confusion with the `clear*` family of methods used for test-only full resets.
- **Unit testing**: Assertion error handling has changed; tests should now expect the new `AssertionError` and use the updated `Assert` methods.

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