import { getAllSnippets, hasTypeScriptCode } from './helpers/test-helpers';
import { compileSnippet } from './helpers/snippet-compiler';

describe('Snippet TypeScript Compilation', () => {
  // Exclude snippets that aren't Office.js snippets
  const EXCLUDED_SNIPPET_IDS = [
    'web-web-default', // Blank web template, not an Office.js snippet
  ];

  const snippets = getAllSnippets()
    .filter(hasTypeScriptCode)
    .filter(snippet => !EXCLUDED_SNIPPET_IDS.includes(snippet.id));

  if (snippets.length === 0) {
    test('No snippets found to test', () => {
      console.warn('Warning: No snippets found in samples/ directory');
      expect(true).toBe(true);
    });
    return;
  }

  test.each(snippets)(
    '$relativePath should compile without errors',
    (snippet) => {
      const result = compileSnippet(snippet);

      if (!result.success) {
        const errorMessages = result.errors.map(err => {
          const location = err.line ? ` at line ${err.line}:${err.column}` : '';
          return `  - ${err.message}${location}`;
        }).join('\n');

        throw new Error(`TypeScript compilation failed for ${snippet.relativePath}:\n${errorMessages}`);
      }

      expect(result.success).toBe(true);
      expect(result.errors).toHaveLength(0);
    }
  );

  // Summary test to report overall stats
  test('Compilation summary', () => {
    const results = snippets.map(snippet => ({
      path: snippet.relativePath,
      result: compileSnippet(snippet)
    }));

    const successful = results.filter(r => r.result.success).length;
    const failed = results.filter(r => !r.result.success).length;

    console.log(`\nCompilation Summary:`);
    console.log(`  Total snippets: ${snippets.length}`);
    console.log(`  Successful: ${successful}`);
    console.log(`  Failed: ${failed}`);

    if (failed > 0) {
      console.log(`\nFailed snippets:`);
      results
        .filter(r => !r.result.success)
        .forEach(r => {
          console.log(`  - ${r.path}: ${r.result.errors.length} error(s)`);
        });
    }

    expect(failed).toBe(0);
  });
});
