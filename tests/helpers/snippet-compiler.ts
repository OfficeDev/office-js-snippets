import * as ts from 'typescript';
import { TestSnippet, usesPreviewApis, getExternalLibraryTypes } from './test-helpers';

export interface CompilationResult {
  success: boolean;
  errors: CompilationError[];
  warnings: string[];
}

export interface CompilationError {
  message: string;
  line?: number;
  column?: number;
  code?: number;
}

/**
 * Compile a snippet's TypeScript code and return compilation results
 */
export function compileSnippet(snippet: TestSnippet): CompilationResult {
  if (!snippet.script || !snippet.script.content) {
    return {
      success: true,
      errors: [],
      warnings: []
    };
  }

  // Wrap snippet code in a compilation harness
  const wrappedCode = createCompilationHarness(snippet);

  // Create a virtual source file
  const sourceFile = ts.createSourceFile(
    'snippet.ts',
    wrappedCode,
    ts.ScriptTarget.ES2024,
    true
  );

  // Determine which Office.js types to use based on the snippet's libraries
  const isPreview = usesPreviewApis(snippet);
  const officeTypes = isPreview ? 'office-js-preview' : 'office-js';

  // Detect external library types (jQuery, Lodash, etc.)
  const externalTypes = getExternalLibraryTypes(snippet);

  // Build complete types array
  const types = [officeTypes, ...externalTypes];

  // Configure compiler options
  const compilerOptions: ts.CompilerOptions = {
    target: ts.ScriptTarget.ES2024,
    module: ts.ModuleKind.CommonJS,
    lib: ['lib.es2024.d.ts', 'lib.dom.d.ts'],
    strict: false,
    noEmit: true,
    skipLibCheck: true,
    skipDefaultLibCheck: true,
    esModuleInterop: true,
    allowSyntheticDefaultImports: true,
    resolveJsonModule: false,
    types: types  // Use appropriate types based on snippet dependencies
  };

  // Create a custom compiler host
  const host = createCompilerHost(sourceFile, compilerOptions);

  // Create program and get diagnostics
  const program = ts.createProgram(['snippet.ts'], compilerOptions, host);
  const diagnostics = ts.getPreEmitDiagnostics(program);

  // Parse diagnostics
  const errors: CompilationError[] = [];
  const warnings: string[] = [];

  for (const diagnostic of diagnostics) {
    // Skip certain non-critical errors
    if (shouldIgnoreDiagnostic(diagnostic)) {
      continue;
    }

    if (diagnostic.category === ts.DiagnosticCategory.Error) {
      const message = ts.flattenDiagnosticMessageText(diagnostic.messageText, '\n');
      const error: CompilationError = {
        message,
        code: diagnostic.code
      };

      if (diagnostic.file) {
        const { line, character } = diagnostic.file.getLineAndCharacterOfPosition(diagnostic.start!);
        error.line = line + 1;
        error.column = character + 1;
      }

      errors.push(error);
    } else if (diagnostic.category === ts.DiagnosticCategory.Warning) {
      warnings.push(ts.flattenDiagnosticMessageText(diagnostic.messageText, '\n'));
    }
  }

  return {
    success: errors.length === 0,
    errors,
    warnings
  };
}

/**
 * Create a compilation harness with necessary declarations and imports
 */
function createCompilationHarness(snippet: TestSnippet): string {
  // We're now using real @types/office-js, so we only need minimal declarations
  // for utilities that aren't part of Office.js types
  let declarations = `
// Utility declarations for helpers that snippets commonly use
interface OfficeHelpers {
  UI: {
    notify(message: string): void;
  };
  Utilities: {
    log(message: string): void;
  };
}
declare const OfficeHelpers: OfficeHelpers | undefined;
`;

  // Add the snippet code
  const snippetCode = snippet.script.content;

  return declarations + '\n\n// Snippet code:\n' + snippetCode;
}

/**
 * Create a custom compiler host that uses our virtual source files
 */
function createCompilerHost(
  sourceFile: ts.SourceFile,
  options: ts.CompilerOptions
): ts.CompilerHost {
  const host = ts.createCompilerHost(options);

  const originalGetSourceFile = host.getSourceFile;
  host.getSourceFile = (fileName, languageVersion) => {
    if (fileName === 'snippet.ts') {
      return sourceFile;
    }
    return originalGetSourceFile(fileName, languageVersion);
  };

  return host;
}

/**
 * Determine if a diagnostic should be ignored
 */
function shouldIgnoreDiagnostic(diagnostic: ts.Diagnostic): boolean {
  // Ignore "Cannot find name" errors for certain globals that may be provided by libraries
  // Note: We no longer ignore $ and jQuery since we include @types for them
  if (diagnostic.code === 2304) { // Cannot find name
    const message = ts.flattenDiagnosticMessageText(diagnostic.messageText, '\n');
    const ignoredNames = ['fabric', 'OfficeHelpers'];
    if (ignoredNames.some(name => message.includes(`'${name}'`))) {
      return true;
    }
  }

  // Ignore module resolution errors for external libraries
  if (diagnostic.code === 2307) { // Cannot find module
    return true;
  }

  // Ignore lodash throttle maxWait error (TS2769 - No overload matches)
  // The maxWait option is not officially documented for _.throttle but works at runtime
  // because throttle is implemented using debounce internally
  if (diagnostic.code === 2769) { // No overload matches this call
    const message = ts.flattenDiagnosticMessageText(diagnostic.messageText, '\n');
    if (message.includes('maxWait') && message.includes('ThrottleSettings')) {
      return true;
    }
  }

  return false;
}
