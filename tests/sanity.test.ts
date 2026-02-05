import { getAllSnippets, getAllSnippetFiles, hasTypeScriptCode } from './helpers/test-helpers';

describe('Test Infrastructure Sanity Check', () => {
  test('should load snippet files', () => {
    const files = getAllSnippetFiles();
    console.log(`Loaded ${files.length} snippet files`);
    expect(files.length).toBeGreaterThan(0);
  });

  test('should parse snippet YAML', () => {
    const snippets = getAllSnippets();
    console.log(`Parsed ${snippets.length} snippets`);
    expect(snippets.length).toBeGreaterThan(0);

    // Check first snippet has required fields
    const first = snippets[0];
    expect(first).toHaveProperty('host');
    expect(first).toHaveProperty('relativePath');
  });

  test('should identify TypeScript snippets', () => {
    const snippets = getAllSnippets();
    const tsSnippets = snippets.filter(hasTypeScriptCode);
    console.log(`Found ${tsSnippets.length} snippets with TypeScript code`);
    expect(tsSnippets.length).toBeGreaterThan(0);
  });

  test('should categorize snippets by host', () => {
    const snippets = getAllSnippets();
    const byHost = new Map<string, number>();

    snippets.forEach(s => {
      const host = s.host ? s.host.toUpperCase() : 'UNKNOWN';
      byHost.set(host, (byHost.get(host) || 0) + 1);
    });

    console.log('\nSnippets by host:');
    byHost.forEach((count, host) => {
      console.log(`  ${host}: ${count}`);
    });

    expect(byHost.size).toBeGreaterThan(0);
  });
});
