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
 */
export async function executeSnippetCode(snippet: TestSnippet): Promise<void> {
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
