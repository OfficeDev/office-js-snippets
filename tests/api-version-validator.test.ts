import { getAllSnippets, hasTypeScriptCode, getApiSetVersion } from './helpers/test-helpers';
import { checkApiVersions, getIncompatibleApis } from './helpers/api-usage-parser';

describe('API Version Compatibility', () => {
  const snippets = getAllSnippets().filter(hasTypeScriptCode);

  if (snippets.length === 0) {
    test('No snippets found to test', () => {
      console.warn('Warning: No snippets found in samples/ directory');
      expect(true).toBe(true);
    });
    return;
  }

  test.each(snippets)(
    '$relativePath should use APIs compatible with declared version',
    (snippet) => {
      const declaredVersion = getApiSetVersion(snippet);

      if (declaredVersion === undefined) {
        // No API set declared, skip validation
        return;
      }

      const incompatibleApis = getIncompatibleApis(snippet);

      if (incompatibleApis.length > 0) {
        const errorMessages = incompatibleApis.map(api => {
          const location = api.line ? ` at line ${api.line}:${api.column}` : '';
          return `  - API '${api.api}' requires version ${api.requiredVersion} but snippet declares ${api.declaredVersion}${location}`;
        }).join('\n');

        fail(`API version mismatch in ${snippet.relativePath}:\n${errorMessages}`);
      }

      expect(incompatibleApis).toHaveLength(0);
    }
  );

  // Summary test
  test('API version validation summary', () => {
    const snippetsWithApiSet = snippets.filter(s => getApiSetVersion(s) !== undefined);
    const results = snippetsWithApiSet.map(snippet => ({
      path: snippet.relativePath,
      host: snippet.host,
      declaredVersion: getApiSetVersion(snippet),
      checks: checkApiVersions(snippet),
      incompatible: getIncompatibleApis(snippet)
    }));

    const totalApis = results.reduce((sum, r) => sum + r.checks.length, 0);
    const totalIncompatible = results.reduce((sum, r) => sum + r.incompatible.length, 0);
    const snippetsWithIssues = results.filter(r => r.incompatible.length > 0);

    console.log(`\nAPI Version Validation Summary:`);
    console.log(`  Total snippets with API set: ${snippetsWithApiSet.length}`);
    console.log(`  Total API calls validated: ${totalApis}`);
    console.log(`  Incompatible API calls: ${totalIncompatible}`);
    console.log(`  Snippets with issues: ${snippetsWithIssues.length}`);

    if (snippetsWithIssues.length > 0) {
      console.log(`\nSnippets with API version issues:`);
      snippetsWithIssues.forEach(r => {
        console.log(`  - ${r.path}:`);
        console.log(`      Declared: ${r.host}Api ${r.declaredVersion}`);
        console.log(`      Incompatible APIs: ${r.incompatible.length}`);
        r.incompatible.slice(0, 3).forEach(api => {
          console.log(`        - ${api.api} (requires ${api.requiredVersion})`);
        });
        if (r.incompatible.length > 3) {
          console.log(`        ... and ${r.incompatible.length - 3} more`);
        }
      });
    }

    // Show version distribution
    const versionCounts = new Map<string, number>();
    results.forEach(r => {
      const key = `${r.host}Api ${r.declaredVersion}`;
      versionCounts.set(key, (versionCounts.get(key) || 0) + 1);
    });

    console.log(`\nAPI Version Distribution:`);
    Array.from(versionCounts.entries())
      .sort((a, b) => a[0].localeCompare(b[0]))
      .forEach(([version, count]) => {
        console.log(`  - ${version}: ${count} snippets`);
      });

    expect(totalIncompatible).toBe(0);
  });
});
