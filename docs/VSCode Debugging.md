# VSCode Debugging with TypeScript

## 1. Terminal (Manual) Debugging

- **Directly run with ts-node:**
  ```sh
  npx ts-node --project tsconfig.test.json wrappers/main-wrapper.ts
  ```
- **To debug:**  
  1. **Set your breakpoints before running.**  
  2. Run with Node.js inspector if you want to attach VSCode:
     ```sh
     node --inspect-brk -r ts-node/register --project tsconfig.test.json wrappers/main-wrapper.ts
     ```
  3. In VSCode, use “Run & Debug” → “Attach to Node Process”.

---

## 2. VSCode Launch Configuration (`.vscode/launch.json`)

- **Allows F5/green arrow debugging from the UI (if it works with your project).**
- **Important:**  
  - **Set breakpoints before pressing F5 or the green arrow.**  
  - For some setups (especially with custom wrappers or initialization), this option may not hit breakpoints unless `preLaunchTask` is set or the command is correct.

**Example launch config:**
```json
{
  "version": "0.2.0",
  "configurations": [
    {
      "type": "node",
      "request": "launch",
      "name": "Debug main-wrapper.ts (ts-node)",
      "runtimeExecutable": "npx",
      "runtimeArgs": [
        "--node-options=--inspect-brk",
        "ts-node",
        "--project",
        "tsconfig.test.json"
      ],
      "args": [
        "${workspaceFolder}/wrappers/main-wrapper.ts"
      ],
      "skipFiles": ["<node_internals>/**"]
    }
  ]
}
```
**How to use:**
1. Open your TypeScript file and **set breakpoints**.
2. Go to Run & Debug panel, select **Debug main-wrapper.ts (ts-node)**.
3. Press F5 / green arrow.

**If breakpoints are not hit:**  
- Try Option 3 (Task + Launch), or use the manual workflow.

---

## 3. VSCode Tasks (`.vscode/tasks.json`) + Launch Configuration

- **Recommended for consistent results when Option 2 fails.**
- **You need both files.**
- **Set breakpoints before running.**

**Example `tasks.json`:**
```json
{
  "version": "2.0.0",
  "tasks": [
    {
      "label": "Run main-wrapper.ts with ts-node",
      "type": "shell",
      "command": "npx ts-node --project tsconfig.test.json wrappers/main-wrapper.ts",
      "group": {
        "kind": "build",
        "isDefault": true
      },
      "presentation": {
        "reveal": "always"
      }
    }
  ]
}
```

**Example `launch.json` (add this configuration):**
```json
{
  "version": "0.2.0",
  "configurations": [
    {
      "type": "node",
      "request": "launch",
      "name": "Debug via Task: main-wrapper.ts (ts-node)",
      "preLaunchTask": "Run main-wrapper.ts with ts-node",
      "runtimeExecutable": "npx",
      "runtimeArgs": [
        "--node-options=--inspect-brk",
        "ts-node",
        "--project",
        "tsconfig.test.json"
      ],
      "args": [
        "${workspaceFolder}/wrappers/main-wrapper.ts"
      ],
      "skipFiles": ["<node_internals>/**"]
    }
  ]
}
```
**How to use:**
1. Open your TypeScript file and **set breakpoints**.
2. Go to Run & Debug, select **Debug via Task: main-wrapper.ts (ts-node)**.
3. Press F5 / green arrow.

---

## 4. Auto Attach in VSCode

- Controls whether VSCode tries to automatically attach the debugger to any Node.js process run from the terminal.
- **Best Practice:** Leave “Auto Attach” off unless you need it.

---

## 5. Troubleshooting Checklist

- Always **set breakpoints before starting** your debug run.
- If breakpoints are not hit:
  - Ensure `"sourceMap": true` is in your `tsconfig`.
  - Delete old `.js`/`.js.map` files.
  - Open VSCode at the project root.
  - Reload VSCode after config changes.
  - Try the Task + Launch workflow if Launch alone does not work.
- For Office Scripts compatibility, use `globalThis` for your `main` function, not `export`.

---

## 6. Quick Reference Table

| Method                | Needs launch.json | Needs tasks.json | Set breakpoints before? | Breakpoints work | Entry Point                       |
|-----------------------|:-----------------:|:----------------:|:----------------------:|:----------------:|-----------------------------------|
| Terminal/CLI (npx)    |        No         |       No         |         Yes            | Yes (if attach)  | wrappers/main-wrapper.ts          |
| VSCode Run/Debug      |       Yes         |       No         |         Yes            | Sometimes*       | wrappers/main-wrapper.ts (config) |
| VSCode Task           |        No         |      Yes         |         Yes            | Yes (if attach)  | wrappers/main-wrapper.ts          |
| Task + Debug (F5)     |       Yes         |      Yes         |         Yes            | Yes              | wrappers/main-wrapper.ts          |

\*Some environments require a task for breakpoints to be hit reliably.

---

**Summary:**  
- **Always set breakpoints before running.**
- If Option 2 (launch config only) doesn’t work, use Option 3 (task + launch config).
- Manual workflow is always available if UI options fail.