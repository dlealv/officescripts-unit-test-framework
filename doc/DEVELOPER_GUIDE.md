# Office Scripts Logging Framework – Development & Testing Workflow

## Overview

This project is a lightweight, extensible logging framework designed for [Office Scripts](https://learn.microsoft.com/en-us/office/dev/scripts/overview/excel) (ExcelScript).  
It is developed in TypeScript to ensure code can run both in the Office Scripts runtime and in local/unit testing scenarios using Node.js, with mock implementations for OfficeScript APIs.

---

## Project Structure

- **`src/`** – Logger source code and framework utilities (production code).
- **`test/`** – Unit tests (entry point: `test/main.ts`).
- **`wrappers/mainWrapper.ts`** – Bootstraps tests using ExcelScript mocks.
- **`mocks/excelscript.mock.ts`** – Local mock implementation of the ExcelScript API.
- **`office-scripts.d.ts`** – Type definitions to enable TS IntelliSense and type safety for Office Scripts.
- **`dist/`** – Production-ready TypeScript source, with test-only code stripped.
- **`package.json`** – Project metadata, scripts, dependencies.
- **`tsconfig.json`** – TypeScript configuration for production build.
- **`tsconfig.test.json`** – TypeScript configuration for local testing. Local tests are executed in a Node.js TypeScript environment using mocks and wrappers that simulate the Office Scripts API. You do not need Excel or Office Online for local testing—just `run npm test` in your terminal
- **`.github/workflows/ci.yml`** – GitHub Actions workflow for Continuous Integration.
- **`eslint.config.js`** – ESLint configuration in ESM format.

## TypeScript Configuration

### Production Build (`tsconfig.json`)

Configured to:
- Use only files from `src/`
- Output to `dist/`
- Exclude test and build output files
  
**Note:**  
> `office-scripts.d.ts` is only required for local development and type-checking. It is not needed in the production build configuration and is not included in the output.

## Development & Testing Workflow

### 1. **Initial Setup**

```sh
# 1. Clone your repository and open in VS Code
git clone <your-repo-url>
cd <your-repo-folder>
code .

# 2. Initialize npm and TypeScript if starting from scratch
npm init -y
npm install typescript ts-node --save-dev
npx tsc --init

# 3. Install dependencies (adjust as needed for your project)
npm install acorn acorn-walk arg create-require diff make-error undici-types v8-compile-cache-lib yn

# 4. Install linting tools and plugins
npm run eslint:setup
```

---

### 2. **Build, Test & Lint Locally**

| **Task**         | **Command**                                 |
|------------------|---------------------------------------------|
| Build project    | `npm run build`                             |
| Run tests        | `npm run test`                              |
| Lint code        | `npm run lint`                              |
| ESLint setup     | `npm run eslint:setup`                      |

#### How it works:
- `npm run build` copies TypeScript source from `src/` to `dist/` and strips test-only code.
- `npm test` runs tests using `ts-node` with `tsconfig.test.json` and the local OfficeScript mocks.
- `npm run lint` lints all `.ts`, `.js`, and `.md` files.
- `npm run eslint:setup` installs/updates ESLint and its plugins.

---

### 3. **Code Management with Git**

| **Task**                  | **Command**                        |
|---------------------------|------------------------------------|
| Stage all changes         | `git add -A`                       |
| Commit changes            | `git commit -m "your message"`     |
| Push to remote            | `git push`                         |
| Pull latest changes       | `git pull`                         |

---

### 4. **Continuous Integration (CI) with GitHub Actions**

To ensure code always builds and tests pass before merging, add the following file:

```yaml name=.github/workflows/ci.yml
name: CI

on:
  push:
    branches: [main]
  pull_request:
    branches: [main]

jobs:
  build-and-test:
    runs-on: ubuntu-latest

    steps:
      - name: Checkout code
        uses: actions/checkout@v4

      - name: Set up Node.js
        uses: actions/setup-node@v4
        with:
          node-version: '20'

      - name: Install dependencies
        run: npm ci

      - name: Build TypeScript
        run: npm run build

      - name: Run tests
        run: npm test
```

**How CI works:**  
- On pushes and pull requests to `main`, GitHub Actions will install dependencies, build your project, and run your tests.
- If anything fails, the workflow fails and merging is blocked (if branch protection is enabled).

**Enforce this workflow:**  
- Go to your GitHub repository → **Settings** → **Branches** → **Branch Protection Rules**.
- Add a rule for `main` and check "Require status checks to pass before merging," selecting your "CI" workflow.

---

## Linting (ESLint)

### Setup

- Run `npm run eslint:setup` to install or update ESLint and plugins:
  - `eslint`
  - `@typescript-eslint/parser`
  - `@typescript-eslint/eslint-plugin`
  - `eslint-plugin-markdown`
- The configuration file is **`eslint.config.js`** in the project root and uses ESM (`import`/`export`) syntax.
- Ensure your `package.json` has `"type": "module"` at the top level (not inside dependencies).

### Usage

- Run `npm run lint` to lint all `.ts`, `.js`, and `.md` files.
- You can customize rules or extend the configuration in `eslint.config.js`.

#### Example ESLint Config

```js
import tseslint from "@typescript-eslint/eslint-plugin";
import tsParser from "@typescript-eslint/parser";

export default [
  {
    files: ["**/*.ts", "**/*.js"],
    languageOptions: { parser: tsParser },
    plugins: { "@typescript-eslint": tseslint },
    rules: { /* Add your rules here */ }
  },
  {
    files: ["**/*.md"],
    plugins: { "markdown": require("eslint-plugin-markdown") }
  }
];
```

### Troubleshooting

- **ESLint ESM Warning:**  
  If you see  
  ```
  Warning: Module type of .../eslint.config.js... is not specified and it doesn't parse as CommonJS...
  ```
  - Ensure `"type": "module"` is present at the root of your `package.json`
  - Try renaming `eslint.config.js` to `eslint.config.mjs` if necessary

- **Plugin Not Found:**  
  ```
  Cannot find package '@typescript-eslint/eslint-plugin'...
  ```
  - Run `npm run eslint:setup`

- If you encounter other issues, try removing `node_modules` and `package-lock.json`, then running `npm install` again.

---

## Scripts Reference

| **Script**         | **Description**                                                      |
|--------------------|----------------------------------------------------------------------|
| `setup`            | Install all dependencies and initialize TypeScript config            |
| `build`            | Copy TypeScript files and strip test-only code for distribution      |
| `copy:ts`          | Copy TypeScript files from `src/` to `dist/`                         |
| `strip:testonly`   | Remove test-only code regions from files in `dist/`                  |
| `test`             | Run tests using ts-node                                              |
| `eslint:setup`     | Install or update ESLint and its plugins                             |
| `lint`             | Lint all `.ts`, `.js`, and `.md` files                               |

---

## Notes on Office Scripts Compatibility

- All code in `src/` must use only APIs available in [Office Scripts](https://learn.microsoft.com/en-us/office/dev/scripts/).
- The local test harness (`wrappers/mainWrapper.ts` and mocks) allows you to run and test code as if it were running in the real Office Scripts environment, but locally, using Node.js.
- Avoid using Node.js or browser-specific APIs in your production (i.e., Office Scripts-targeted) code.

---

## Typical Development Flow

1. **Edit source or test files** in VS Code.
2. **Build and test locally** with:
   ```sh
   npm run build
   npm test
   ```
3. **Stage, commit, and push:**
   ```sh
   git add -A
   git commit -m "Describe your change"
   git push
   ```
4. **Create a pull request** (PR).
5. **CI will run build and test automatically.**
6. **Merge only if checks pass.**

---

## Useful VS Code Terminal Commands

| **Purpose**                | **Command**                                                    |
|----------------------------|---------------------------------------------------------------|
| Initialize npm project     | `npm init` or `npm init -y`                                   |
| Install dependencies       | `npm install` or `npm ci`                                     |
| Add a dependency           | `npm install <package>`                                       |
| Add a dev dependency       | `npm install --save-dev <package>`                            |
| Build TypeScript           | `npm run build`                                               |
| Run tests                  | `npm test`                                                    |
| Lint code                  | `npm run lint`                                                |
| Stage all changes          | `git add -A`                                                  |
| Commit changes             | `git commit -m "your message"`                                |
| Push to GitHub             | `git push`                                                    |
| Pull from GitHub           | `git pull`                                                    |
| Open VS Code               | `code .`                                                      |
| Install TypeScript         | `npm install typescript --save-dev`                           |
| Install ts-node            | `npm install ts-node --save-dev`                              |
| Initialize tsconfig        | `npx tsc --init`                                              |

---

## Troubleshooting

- If you encounter type or runtime errors, double-check your mocks and type declarations to ensure compatibility with the real Office Scripts API.
- For new test files, import them in `test/main.ts` or ensure your test runner discovers them.
- For ESLint issues, see the "Linting (ESLint)" section above.

---

## Further Improvements

- Add code coverage tools for stricter code quality.
- Enhance your mock OfficeScript API as needed for more advanced scenarios.
- Keep your README and this guide updated as your workflow evolves!

---

**Happy scripting and testing!**
