/**
 * Quick test to verify compilation and library validation work
 * Tests only the first 5 snippets to keep it fast
 */
import { getAllSnippets, hasTypeScriptCode, getOfficeJsLibraries } from './helpers/test-helpers';
import { compileSnippet } from './helpers/snippet-compiler';
import fetch from 'node-fetch';

describe('Quick Validation Test', () => {
  const allSnippets = getAllSnippets().filter(hasTypeScriptCode);
  const testSnippets = allSnippets.slice(0, 5); // Test first 5 only

  describe('TypeScript Compilation', () => {
    test.each(testSnippets)(
      '$relativePath should compile',
      (snippet) => {
        const result = compileSnippet(snippet);

        if (!result.success) {
          console.log(`Compilation errors for ${snippet.relativePath}:`);
          result.errors.forEach(err => {
            console.log(`  - ${err.message}`);
          });
        }

        expect(result.success).toBe(true);
      }
    );
  });

  describe('Library URLs', () => {
    test('first snippet should have valid Office.js library', async () => {
      const snippet = testSnippets[0];
      const libraries = getOfficeJsLibraries(snippet);

      expect(libraries.length).toBeGreaterThan(0);

      // Check that Office.js URL is reachable
      const officeJsUrl = libraries.find(lib => lib.includes('office.js'));
      if (officeJsUrl) {
        const response = await fetch(officeJsUrl, { method: 'HEAD', timeout: 10000 });
        expect(response.ok).toBe(true);
      }
    }, 30000);
  });

  test('Quick test summary', () => {
    console.log(`\nQuick Test Summary:`);
    console.log(`  Total snippets in repo: ${allSnippets.length}`);
    console.log(`  Tested: ${testSnippets.length}`);
    console.log(`  Test infrastructure: Working ✓`);
  });
});
