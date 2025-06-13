# Office Scripts Logging Framework – Development & Testing Workflow

## Overview

This project is a lightweight, extensible logging framework designed for [Office Scripts](https://learn.microsoft.com/en-us/office/dev/scripts/overview/excel) (ExcelScript).  
It is developed in TypeScript to ensure code can run both in the Office Scripts runtime and in local/unit testing scenarios using Node.js, with mock implementations for OfficeScript APIs.

---

## Project Structure

- **`src/`** – Logger source code and framework utilities (production code).
- **`test/`** – Unit tests (entry point: `test/main.ts`).
- **`wrappers/main-wrapper.ts`** – Bootstraps tests using ExcelScript mocks.
- **`mocks/excelscript.mock.ts`** – Local mock implementation of the ExcelScript API.
- **`office-scripts.d.ts`** – Type definitions to enable TS IntelliSense and type safety for Office Scripts.
- **`dist/`** – Production-ready TypeScript source, with test-only code stripped.
- **`package.json`** – Project metadata, scripts, dependencies.
- **`tsconfig.json`** – TypeScript configuration for production build.
- **`tsconfig.test.json`** – TypeScript configuration for local testing. Local tests are executed in a Node.js TypeScript environment using mocks and wrappers that simulate the Office Scripts API.
- **`.github/workflows/ci.yml`** – GitHub Actions workflow for Continuous Integration.
- **`.vscode/`** – (Optional) VS Code workspace settings and launch configurations.
- **`node_modules/`** – Installed project dependencies (auto-generated, not committed).
- **`.gitignore`** – Specifies files and folders for Git to ignore (e.g., `node_modules/`, `dist/`, etc.).

## TypeScript Configuration

<!-- ...rest of your document remains unchanged... -->
