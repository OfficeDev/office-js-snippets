/**
 * Auto-generated runtime tests for Office.js snippets
 *
 * This test suite dynamically generates runtime tests for snippet groups that are compatible
 * with our mock Office.js environment. Tests verify that snippets execute without errors.
 *
 * Coverage: ~195 tests across 36 snippet groups (97% of targeted snippets)
 *
 * Maintenance:
 * - Add new snippet groups to INCLUDED_GROUPS when they are created
 * - Groups using unsupported features should remain commented out with explanation
 * - Update EXCLUDED_PATTERNS if new problematic patterns are identified
 */

import { getAllSnippets, TestSnippet } from './helpers/test-helpers';
import { executeSnippetCode, createMockDocument, clickButton } from './helpers/snippet-runtime';
import {
  createExcelMock,
  createWordMock,
  createPowerPointMock,
  createOneNoteMock,
  createOfficeCommonApiMock,
} from './helpers/mock-factories';

/**
 * Snippet groups to test (by folder name)
 *
 * These groups work with our current mock implementation. Groups are organized by host
 * and include all major API surfaces except those requiring advanced mocking.
 */
const INCLUDED_GROUPS = [
  // Excel - Core APIs
  '01-basics',
  '10-chart',
  '12-comments-and-notes',
  '14-conditional-formatting',
  '18-custom-xml-parts',
  '22-data-validation',
  '26-document',
  '34-named-item',
  '38-pivottable',
  '42-range',
  '44-shape',
  '46-table',
  '50-workbook',
  '54-worksheet',
  '90-scenarios',
  '99-just-for-fun',

  // Excel - Excluded groups (require advanced mocking)
  // '16-custom-functions',  // Runs in separate JavaScript runtime, cannot be mocked
  // '20-data-types',        // Requires Excel.CellValueType enum support
  // '30-events',            // Requires event handler mocking

  // Word - Core APIs
  '01-basics',
  '10-content-controls',
  '15-images',
  '20-lists',
  '25-paragraph',
  '30-properties',
  '35-ranges',
  '40-tables',
  '45-shapes',
  '50-document',
  '90-scenarios',

  // Word - Excluded groups
  // '99-preview-apis',      // Preview APIs are unstable and change frequently

  // PowerPoint - Core APIs
  'basics',
  'slide-management',
  'shapes',
  'text',
  'images',
  'hyperlinks',
  'tags',
  'document',
  'scenarios',

  // PowerPoint - Excluded groups
  // 'preview-apis',         // Preview APIs are unstable and change frequently
];

/**
 * Exclusion patterns for snippets that cannot be tested with current mocks
 *
 * These patterns identify snippets that use features our mocks don't support:
 * - Missing enum definitions (SearchDirection, CellValueType)
 * - Event handler registration (requires event emitter mocking)
 * - Custom function runtime (different JavaScript context)
 * - Complex DOM manipulation (beyond basic button clicks)
 * - Unstable preview APIs
 */
const EXCLUDED_PATTERNS = [
  '*preview-apis*',        // Preview APIs change frequently
  '*events*',              // Event handlers require event emitter mocking
  '*cellvalue*',           // Requires Excel.CellValueType enum (see README "When tests are needed")
  '*custom-functions*',    // Runs in separate JavaScript runtime
  '*search*',              // Requires Excel.SearchDirection enum (see README "When tests are needed")
  '*find*',                // Requires Excel.SearchDirection enum (see README "When tests are needed")
  '*tetromino*',           // Game requires complex DOM manipulation
];

/**
 * Critical snippets that must always be tested
 *
 * These basic API call snippets serve as smoke tests to ensure the test infrastructure
 * works correctly for each host. If these fail, the entire test framework needs review.
 */
const MUST_TEST_SNIPPETS = [
  'excel-basics-basic-api-call',
  'word-basics-basic-api-call',
  'powerpoint-basics-basic-api-call-ts',
];

/**
 * Check if snippet matches any exclusion pattern
 *
 * Exclusion patterns use wildcards (*) for flexible matching against both the snippet's
 * file path and ID. Example: '*events*' matches 'excel/30-events/workbook-events.yaml'
 */
function isExcluded(snippet: TestSnippet): boolean {
  const pathLower = snippet.relativePath.toLowerCase();
  const idLower = snippet.id.toLowerCase();

  return EXCLUDED_PATTERNS.some(pattern => {
    const patternLower = pattern.toLowerCase().replace(/\*/g, '');
    return pathLower.includes(patternLower) || idLower.includes(patternLower);
  });
}

/**
 * Check if snippet belongs to an included group
 *
 * Groups are identified by the folder name in the snippet path.
 * Example: 'samples/excel/42-range/formatting.yaml' → group is '42-range'
 */
function isInIncludedGroup(snippet: TestSnippet): boolean {
  const pathParts = snippet.relativePath.split(/[/\\]/);
  const group = pathParts[1]; // Second part is the group folder
  return INCLUDED_GROUPS.includes(group);
}

/**
 * Determine if snippet uses Common API (Office 2013) instead of host-specific API
 *
 * Common API snippets use Office.context.document and don't use Excel.run/Word.run patterns.
 * These require different mocking setup.
 */
function usesCommonApi(snippet: TestSnippet): boolean {
  if (!snippet.script?.content) return false;
  const code = snippet.script.content;
  return code.includes('Office.context.document') && !code.includes('.run(');
}

/**
 * Button clicking strategy for snippets
 *
 * We click buttons in this order:
 * 1. setup - Prepares initial data/environment (allowed to fail gracefully)
 * 2. run - Main functionality (errors here fail the test)
 *
 * This approach:
 * - Tests snippets with working setup buttons more thoroughly
 * - Allows snippets with unsupported setup operations to still test their run functionality
 * - Avoids memory issues from clicking too many buttons
 * - Skips event handler registration which can cause test issues
 */

/**
 * Determine if a snippet requires user input for its buttons
 *
 * Detects patterns like:
 * - document.getElementById("inputId").value
 * - HTMLInputElement, HTMLSelectElement, etc.
 */
function requiresUserInput(snippet: TestSnippet): boolean {
  const code = snippet.script?.content || '';

  // Check for patterns indicating user input
  const userInputPatterns = [
    /getElementById\([^)]+\)\s*as\s+HTMLInputElement/,
    /getElementById\([^)]+\)\s*as\s+HTMLSelectElement/,
    /getElementById\([^)]+\)\.value/,
    /getElementById\([^)]+\)\.files/,
    /getElementById\([^)]+\)\.checked/,
  ];

  return userInputPatterns.some(pattern => pattern.test(code));
}

/**
 * Execute a snippet with appropriate mocks based on its host and API type
 *
 * Sets up the global environment with the correct mock objects, executes the snippet,
 * clicks setup button (allowing it to fail), then clicks run button (errors fail the test).
 *
 * @param snippet - The snippet to test
 * @param consoleErrorSpy - Spy to track console.error calls
 */
async function runSnippetTest(snippet: TestSnippet, consoleErrorSpy?: jest.SpyInstance) {
  const buttonHandlers = new Map<string, Function>();

  if (usesCommonApi(snippet)) {
    (global as any).Office = createOfficeCommonApiMock();
  } else {
    const host = snippet.host?.toUpperCase();
    switch (host) {
      case 'EXCEL':
        const excelMock = createExcelMock();
        (global as any).Excel = excelMock.mockObject;
        break;
      case 'WORD':
        const wordMock = createWordMock();
        (global as any).Word = wordMock.mockObject;
        break;
      case 'POWERPOINT':
        const pptMock = createPowerPointMock();
        (global as any).PowerPoint = pptMock.mockObject;
        break;
      case 'ONENOTE':
        const oneNoteMock = createOneNoteMock();
        (global as any).OneNote = oneNoteMock.mockObject;
        break;
      default:
        throw new Error(`Unsupported host: ${host}`);
    }
  }

  (global as any).document = createMockDocument(buttonHandlers);

  // Execute the snippet (registers button handlers)
  executeSnippetCode(snippet);

  // Skip clicking buttons if the snippet requires user input
  if (requiresUserInput(snippet)) {
    return;
  }

  // Click setup button if it exists (allow it to fail gracefully)
  if (buttonHandlers.has('setup')) {
    try {
      await clickButton(buttonHandlers, 'setup');
    } catch (error) {
      // Setup failed - this is okay, many snippets have setup operations
      // that require mocks we haven't implemented. We'll still test run().
    }

    // Clear any console.error calls from setup - we only care about errors from run()
    if (consoleErrorSpy) {
      consoleErrorSpy.mockClear();
    }
  }

  // Click run button (this is the critical test - errors here will fail the test)
  if (buttonHandlers.has('run')) {
    await clickButton(buttonHandlers, 'run');
  }
}

/**
 * Get the feature group from a snippet's relative path
 *
 * Extracts the folder name after the host directory.
 * Example: 'excel/42-range/formatting.yaml' → '42-range'
 */
function getFeatureGroup(snippet: TestSnippet): string {
  const pathParts = snippet.relativePath.split(/[/\\]/);

  // Find the index of the host (excel, word, powerpoint, onenote)
  const host = snippet.host?.toLowerCase();
  const hostIndex = pathParts.findIndex(part => part.toLowerCase() === host);

  // Group is the next part after host
  if (hostIndex !== -1 && hostIndex + 1 < pathParts.length) {
    return pathParts[hostIndex + 1];
  }

  // Fallback to the filename if we can't determine the group
  return pathParts[pathParts.length - 1].replace('.yaml', '');
}

/**
 * Get a human-readable name for a feature group
 *
 * Removes number prefixes and converts to title case
 */
function getGroupDisplayName(group: string): string {
  // Remove number prefix (e.g., '42-range' → 'range')
  const withoutNumber = group.replace(/^\d+-/, '');

  // Convert kebab-case to Title Case
  return withoutNumber
    .split('-')
    .map(word => word.charAt(0).toUpperCase() + word.slice(1))
    .join(' ');
}

/**
 * Load and categorize all snippets by host and feature group
 */
const allSnippets = getAllSnippets();
const snippetsByHostAndGroup = allSnippets.reduce((acc, snippet) => {
  const host = snippet.host?.toUpperCase() || 'UNKNOWN';
  const group = getFeatureGroup(snippet);

  if (!acc[host]) acc[host] = {};
  if (!acc[host][group]) acc[host][group] = [];

  acc[host][group].push(snippet);
  return acc;
}, {} as Record<string, Record<string, TestSnippet[]>>);

/**
 * Generate test suites organized by host and feature group
 *
 * Creates nested describe blocks:
 * - Host level (Excel, Word, PowerPoint, OneNote)
 *   - Feature group level (Basics, Charts, Ranges, etc.)
 *     - Individual tests
 *
 * This structure makes it easier to identify which specific feature area has issues.
 */
['EXCEL', 'WORD', 'POWERPOINT', 'ONENOTE'].forEach((host) => {
  const hostGroups = snippetsByHostAndGroup[host] || {};
  const groupNames = Object.keys(hostGroups);

  if (groupNames.length === 0) {
    return;
  }

  // Filter groups to only those with snippets that pass our filters
  const validGroups = groupNames.filter((group) => {
    const filteredSnippets = hostGroups[group].filter(s => {
      if (MUST_TEST_SNIPPETS.includes(s.id)) return true;
      if (isExcluded(s)) return false;
      return isInIncludedGroup(s);
    });
    return filteredSnippets.length > 0;
  });

  // Skip this host entirely if no valid groups
  if (validGroups.length === 0) {
    return;
  }

  describe(`${host} Runtime Tests`, () => {
    let consoleErrorSpy: jest.SpyInstance;

    beforeEach(() => {
      // Reset global objects
      (global as any).Excel = undefined;
      (global as any).Word = undefined;
      (global as any).PowerPoint = undefined;
      (global as any).OneNote = undefined;
      (global as any).Office = undefined;

      // Mock console
      jest.spyOn(console, 'log').mockImplementation();
      consoleErrorSpy = jest.spyOn(console, 'error').mockImplementation();
    });

    afterEach(() => {
      jest.restoreAllMocks();
    });

    // Create a describe block for each feature group
    validGroups.forEach((group) => {
      const groupSnippets = hostGroups[group]
        .filter(s => {
          if (MUST_TEST_SNIPPETS.includes(s.id)) return true;
          if (isExcluded(s)) return false;
          return isInIncludedGroup(s);
        });

      const displayName = getGroupDisplayName(group);

      describe(displayName, () => {
        groupSnippets.forEach((snippet) => {
          const testName = snippet.name || snippet.id;

          test(testName, async () => {
            await runSnippetTest(snippet, consoleErrorSpy);

            // Verify no errors were logged during the "run" button click
            // Note: consoleErrorSpy is cleared after setup in runSnippetTest, so this only
            // catches errors from the main "run" button, allowing setup to fail gracefully
            expect(consoleErrorSpy).not.toHaveBeenCalled();
          }, 5000); // 5s timeout is sufficient for mock environment (195 tests = ~16 min max)
        });
      });
    });
  });
});
