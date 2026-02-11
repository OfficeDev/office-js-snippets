import { getAllSnippets, getOfficeJsLibraries, TestSnippet } from './helpers/test-helpers';

// Cache for URL validation results to avoid repeated requests
const urlCache = new Map<string, { ok: boolean; status: number }>();

// URLs to skip (e.g., type definitions that aren't actual URLs)
const SKIP_PATTERNS = [
  /^@types\//,  // npm @types packages
  /^dt~/,       // DefinitelyTyped notation
];

/**
 * Check if a URL should be validated
 */
function shouldValidateUrl(url: string): boolean {
  return SKIP_PATTERNS.every(pattern => !pattern.test(url));
}

/**
 * Validate a single URL with caching and retry logic
 */
async function validateUrl(url: string, retries = 2): Promise<{ ok: boolean; status: number; error?: string }> {
  // Check cache first
  if (urlCache.has(url)) {
    return urlCache.get(url)!;
  }

  for (let attempt = 0; attempt <= retries; attempt++) {
    try {
      // Use AbortController for timeout with native fetch
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 10000); // 10 second timeout

      const response = await fetch(url, {
        method: 'HEAD',
        signal: controller.signal,
        redirect: 'follow'
      });

      clearTimeout(timeoutId);

      const result = {
        ok: response.ok,
        status: response.status
      };

      // Cache successful results and 404s (persistent failures)
      if (response.ok || response.status === 404) {
        urlCache.set(url, result);
      }

      return result;
    } catch (error) {
      // If this is the last retry, return the error
      if (attempt === retries) {
        const result = {
          ok: false,
          status: 0,
          error: error.message || String(error)
        };
        return result;
      }

      // Wait before retrying (exponential backoff)
      await new Promise(resolve => setTimeout(resolve, 1000 * (attempt + 1)));
    }
  }

  return { ok: false, status: 0, error: 'Max retries exceeded' };
}

/**
 * Validate all libraries for a snippet
 */
async function validateSnippetLibraries(snippet: TestSnippet): Promise<{ url: string; ok: boolean; status: number; error?: string }[]> {
  const libraries = getOfficeJsLibraries(snippet);
  const results: { url: string; ok: boolean; status: number; error?: string }[] = [];

  for (const url of libraries) {
    if (!shouldValidateUrl(url)) {
      continue; // Skip non-HTTP URLs
    }

    const result = await validateUrl(url);
    results.push({ url, ...result });
  }

  return results;
}

describe('Library URL Validation', () => {
  // Exclude snippets that aren't Office.js snippets
  const EXCLUDED_SNIPPET_IDS = [
    'web-web-default', // Blank web template, not an Office.js snippet
  ];

  const snippets = getAllSnippets()
    .filter(snippet => !EXCLUDED_SNIPPET_IDS.includes(snippet.id));

  if (snippets.length === 0) {
    test('No snippets found to test', () => {
      console.warn('Warning: No snippets found in samples/ directory');
      expect(true).toBe(true);
    });
    return;
  }

  // Group tests by snippet to improve readability
  describe.each(snippets)('$relativePath', (snippet) => {
    test('should have reachable library URLs', async () => {
      const libraries = getOfficeJsLibraries(snippet);

      if (libraries.length === 0) {
        console.warn(`No libraries found for ${snippet.relativePath}`);
        return;
      }

      const validatableLibraries = libraries.filter(shouldValidateUrl);

      if (validatableLibraries.length === 0) {
        // All libraries are @types or other non-HTTP references
        return;
      }

      const results = await validateSnippetLibraries(snippet);
      const failures = results.filter(r => !r.ok);

      if (failures.length > 0) {
        const errorMessages = failures.map(f => {
          const statusMsg = f.status > 0 ? ` (HTTP ${f.status})` : '';
          const errorMsg = f.error ? `: ${f.error}` : '';
          return `  - ${f.url}${statusMsg}${errorMsg}`;
        }).join('\n');

        fail(`Library validation failed for ${snippet.relativePath}:\n${errorMessages}`);
      }

      expect(failures).toHaveLength(0);
    }, 30000); // 30 second timeout per snippet
  });

  // Summary test
  test('Library validation summary', async () => {
    const allLibraries = new Set<string>();
    const failedUrls = new Map<string, string[]>(); // url -> list of snippets using it

    for (const snippet of snippets) {
      const libraries = getOfficeJsLibraries(snippet).filter(shouldValidateUrl);

      for (const url of libraries) {
        allLibraries.add(url);
      }
    }

    console.log(`\nLibrary Validation Summary:`);
    console.log(`  Total unique libraries: ${allLibraries.size}`);
    console.log(`  Total snippets checked: ${snippets.length}`);

    // Validate all unique URLs
    for (const url of allLibraries) {
      const result = await validateUrl(url);
      if (!result.ok) {
        const snippetsUsingUrl = snippets
          .filter(s => getOfficeJsLibraries(s).includes(url))
          .map(s => s.relativePath);
        failedUrls.set(url, snippetsUsingUrl);
      }
    }

    if (failedUrls.size > 0) {
      console.log(`\nFailed URLs:`);
      failedUrls.forEach((snippetPaths, url) => {
        console.log(`  - ${url}`);
        console.log(`    Used in ${snippetPaths.length} snippet(s):`);
        snippetPaths.slice(0, 3).forEach(path => {
          console.log(`      - ${path}`);
        });
        if (snippetPaths.length > 3) {
          console.log(`      ... and ${snippetPaths.length - 3} more`);
        }
      });
    }

    console.log(`  Cache size: ${urlCache.size} URLs cached`);

    expect(failedUrls.size).toBe(0);
  }, 60000); // 60 second timeout for summary
});
