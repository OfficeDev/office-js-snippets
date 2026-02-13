import { TestSnippet } from './test-helpers';
import * as ts from 'typescript';

/**
 * Transpile TypeScript code to JavaScript
 */
function transpileTypeScript(code: string): string {
  const result = ts.transpileModule(code, {
    compilerOptions: {
      target: ts.ScriptTarget.ES2024,
      module: ts.ModuleKind.None,
      removeComments: false,
    },
  });
  return result.outputText;
}

/**
 * Execute a snippet's code in a controlled runtime environment
 *
 * TESTING SCOPE: This executes snippet code for SYNTAX VERIFICATION only.
 * It does NOT verify correct Office.js behavior or real-world functionality.
 *
 * What this tests:
 * ✅ Code executes without JavaScript syntax errors
 * ✅ Button handlers register correctly
 * ✅ No undefined variable references or typos
 *
 * What this does NOT test:
 * ❌ Correct behavior with real Office collections
 * ❌ Proper load/sync patterns
 * ❌ Visual output or document state
 * ❌ Error handling with real Office.js
 *
 * Implementation notes:
 * - Synchronous execution - runs initialization code immediately
 * - Button handlers registered but not yet clicked
 * - Async operations executed separately via clickButton()
 * - Global mocks (Excel, Word, etc.) already set up before execution
 *
 * Security: Uses Function constructor to execute snippet code. This is safe because:
 * - All snippets are from this repository (trusted source)
 * - Code is validated during build process (npm start)
 * - Code is reviewed in PRs before merging
 * - Tests run in isolated Node.js environment, not in browser
 * - No user-provided or external code is executed
 */
export function executeSnippetCode(snippet: TestSnippet): void {
  const code = snippet.script.content;

  // Transpile TypeScript to JavaScript if needed
  const jsCode = snippet.script.language === 'typescript'
    ? transpileTypeScript(code)
    : code;

  // Use Function constructor to create a function from the code
  // This allows us to execute the code in the current scope where global mocks are set up
  const func = new Function(jsCode);
  func();
}

/**
 * Create a mock DOM button element that can trigger snippet execution
 */
export function createMockButton(handler: Function) {
  return {
    addEventListener: (event: string, callback: Function) => {
      if (event === 'click') {
        // Store the callback so we can trigger it later
        handler(callback);
      }
    },
  };
}

/**
 * Create a mock document with getElementById support
 */
export function createMockDocument(buttonHandlers: Map<string, Function>) {
  return {
    getElementById: (id: string) => {
      return {
        addEventListener: (event: string, callback: Function) => {
          if (event === 'click') {
            // Store the handler so tests can trigger it
            buttonHandlers.set(id, callback);
          }
        },
      };
    },
  };
}

/**
 * Trigger a button click event
 */
export async function clickButton(buttonHandlers: Map<string, Function>, buttonId: string): Promise<void> {
  const handler = buttonHandlers.get(buttonId);
  if (!handler) {
    throw new Error(`No handler found for button: ${buttonId}`);
  }
  await handler();
}
