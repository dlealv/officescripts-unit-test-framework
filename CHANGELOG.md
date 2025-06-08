# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [Unreleased]

### Added
- Initial `CHANGELOG.md` to document project changes.
- Additional ESLint configuration and troubleshooting details in the Developer Guide.
- `eslint.config.js` in ESM format, with corresponding `package.json` `"type": "module"` at the root.
- New script commands:
  - `eslint:setup` – Installs/updates ESLint and plugins.
  - Updated `lint` – Lints `.ts`, `.js`, and `.md` files.
  - `copy:ts` and `strip:testonly` for improved build process.

### Changed
- Developer Guide now includes explicit project structure, more detailed linting section, and updated scripts reference.

### Fixed
- Fixed issue where `"type": "module"` was incorrectly listed as a dependency.

---

## [1.0.0] – 2025-06-08

### Added
- Initial release of the Office Scripts Logging Framework.
- TypeScript-based logger for Office Scripts with mock/test harness.
- GitHub Actions CI workflow.
- Developer workflow and project documentation.

---
