// src/wrapper.ts

// Import Office Script file as side-effect, so it registers globals like main(), TestCase
import './main.ts';

// Now explicitly export them from global scope for Node.js tests
export const mainFn = (globalThis as any).main;
export const TestCaseClass = (globalThis as any).TestCase;

