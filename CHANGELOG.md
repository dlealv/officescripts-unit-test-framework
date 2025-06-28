# Changelog

All notable changes to this project will be documented in this file.  
This project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2025-06-29

### Added
- Initial release of the Office Scripts Unit Testing Framework.
- Provides `Assert` class with static assertion methods for:
  - Value equality and inequality: `equals`, `notEquals` (supports primitives, arrays, and deep object checks)
  - Null and undefined checks: `isNull`, `isNotNull`, `isUndefined`, `isNotUndefined`, `isDefined`
  - Truthiness checks: `isTrue`, `isFalse`
  - Type checking: `isType`, `isNotType`
  - Instance checks: `isInstanceOf`, `isNotInstanceOf`
  - Containment: `contains` (for arrays and strings)
  - Exception assertions: `throws`, `doesNotThrow`
  - Manual failure: `fail`
- Provides `TestRunner` class for structured test execution and configurable verbosity (`OFF`, `HEADER`, `SECTION`, `SUBSECTION`).
- Office Scripts and Node/TypeScript compatibility.
- Example test suite and usage documentation.