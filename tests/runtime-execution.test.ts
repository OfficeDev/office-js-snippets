import {
  runExcelSnippetTest,
  runWordSnippetTest,
  runPowerPointSnippetTest,
  runCommonApiSnippetTest,
} from './helpers/snippet-test-runner';
import * as path from 'path';
import { createWordMock } from './helpers/mock-factories';
import { loadSnippetByPath } from './helpers/test-helpers';
import { executeSnippetCode, createMockDocument, clickButton } from './helpers/snippet-runtime';

describe('Runtime Execution Tests - Excel', () => {
  let consoleSpy: jest.SpyInstance;
  let consoleErrorSpy: jest.SpyInstance;

  beforeEach(() => {
    (global as any).Excel = undefined;
    (global as any).Office = undefined;
    consoleSpy = jest.spyOn(console, 'log').mockImplementation(() => {});
    consoleErrorSpy = jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  test('Excel: basic-api-call executes without runtime errors', async () => {
    await runExcelSnippetTest({
      snippetPath: path.join('samples', 'excel', '01-basics', 'basic-api-call.yaml'),
      assertions: ({ mockContext, mockRange }) => {
        // Verify Excel.run was called
        expect(mockContext).toBeDefined();

        // Verify range operations
        expect(mockContext.workbook.getSelectedRange).toHaveBeenCalled();
        expect(mockRange.load).toHaveBeenCalledWith('address');
        expect(mockRange.format.fill.color).toBe('yellow');
        expect(mockContext.sync).toHaveBeenCalled();

        // Verify console output
        expect(consoleSpy).toHaveBeenCalledWith(expect.stringContaining('The range address was'));

        // Verify no errors were logged
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test('Excel: basic-common-api-call executes without runtime errors', async () => {
    await runCommonApiSnippetTest({
      snippetPath: path.join('samples', 'excel', '01-basics', 'basic-common-api-call.yaml'),
      assertions: ({ Office }) => {
        // Verify Common API was called
        expect(Office.context.document.getSelectedDataAsync).toHaveBeenCalled();

        // Verify console output
        expect(consoleSpy).toHaveBeenCalledWith(expect.stringContaining('The selected data is'));

        // Verify no errors were logged
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    'setup',
    'set-range-font',
    'get-range-font',
    'set-fill-color',
    'set-number-format',
  ])('Excel: range formatting %s action executes without runtime errors', async (buttonId) => {
    await runExcelSnippetTest({
      snippetPath: path.join('samples', 'excel', '42-range', 'formatting.yaml'),
      buttonId,
      assertions: ({ mockContext }) => {
        expect(mockContext.sync).toHaveBeenCalled();
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    'setup',
    'add-rows',
    'get-row-count',
    'get-first-row',
    'delete-rows',
  ])('Excel: manage table rows %s action executes without runtime errors', async (buttonId) => {
    await runExcelSnippetTest({
      snippetPath: path.join('samples', 'excel', '46-table', 'manage-table-rows.yaml'),
      buttonId,
    });

    expect(consoleErrorSpy).not.toHaveBeenCalled();
  });
});

describe('Runtime Execution Tests - Word', () => {
  let consoleSpy: jest.SpyInstance;
  let consoleErrorSpy: jest.SpyInstance;

  beforeEach(() => {
    (global as any).Word = undefined;
    (global as any).Office = undefined;
    consoleSpy = jest.spyOn(console, 'log').mockImplementation(() => {});
    consoleErrorSpy = jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  test('Word: basic-api-call executes without runtime errors', async () => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '01-basics', 'basic-api-call.yaml'),
      assertions: ({ mockContext, mockRange }) => {
        // Verify Word.run was called
        expect(mockContext).toBeDefined();

        // Verify range operations
        expect(mockContext.document.getSelection).toHaveBeenCalled();
        expect(mockRange.font.color).toBe('red');
        expect(mockRange.load).toHaveBeenCalledWith('text');
        expect(mockContext.sync).toHaveBeenCalled();

        // Verify console output
        expect(consoleSpy).toHaveBeenCalledWith(expect.stringContaining('The selected text was'));

        // Verify no errors were logged
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test('Word: basic-common-api-call executes without runtime errors', async () => {
    await runCommonApiSnippetTest({
      snippetPath: path.join('samples', 'word', '01-basics', 'basic-common-api-call.yaml'),
      assertions: ({ Office }) => {
        // Verify Common API was called
        expect(Office.context.document.getSelectedDataAsync).toHaveBeenCalled();

        // Verify console output
        expect(consoleSpy).toHaveBeenCalledWith(expect.stringContaining('The selected data is'));

        // Verify no errors were logged
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test('Word: get document breaks executes without runtime errors', async () => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '35-ranges', 'get-pages.yaml'),
      buttonId: 'get-breaks',
      assertions: ({ mockContext }) => {
        const breaks = mockContext.document.activeWindow.activePane.pages.items[0].breaks;
        expect(breaks.load).toHaveBeenCalledWith('items/pageIndex');
        expect(consoleSpy).toHaveBeenCalledWith('Breaks found: 1');
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  const rangeScenarios = [
    ['35-ranges', 'basic-range-operations.yaml', [
      'setup', 'get-range', 'replace-text', 'insert-paragraph', 'insert-break',
      'search-range', 'set-style', 'clear-range', 'delete-range',
    ]],
    ['35-ranges', 'compare-location.yaml', [
      'setup', 'compare', 'compare-to-selection', 'expand-ranges', 'expand-ranges-safely',
      'intersect-ranges', 'intersect-ranges-safely',
    ]],
    ['35-ranges', 'split-words-of-first-paragraph.yaml', [
      'setup', 'highlight', 'get-next-sentence', 'get-next-sentence-safely',
    ]],
    ['35-ranges', 'scroll-to-range.yaml', [
      'setup', 'scroll', 'scroll-end', 'highlight-range', 'remove-highlight',
    ]],
    ['35-ranges', 'insert-and-read-formatted-content.yaml', [
      'setup', 'insert-html', 'insert-ooxml', 'insert-file', 'get-html', 'get-ooxml',
    ]],
    ['35-ranges', 'manage-range-bookmarks.yaml', [
      'setup', 'insert-bookmark', 'get-bookmark-names', 'get-bookmarks',
    ]],
    ['35-ranges', 'inspect-range-structure.yaml', [
      'setup', 'get-controls', 'get-fields', 'get-pictures', 'get-lists', 'insert-table',
      'get-tables', 'get-parent-table', 'get-parent-cell', 'check-parent-table',
      'check-parent-cell', 'set-hyperlink', 'get-hyperlink-ranges',
    ]],
    ['35-ranges', 'format-range-text-desktop.yaml', [
      'setup', 'get-formatting', 'get-position', 'move-start', 'move-end', 'set-case',
      'set-width', 'combine-characters', 'fit-text', 'set-horizontal-in-vertical',
      'set-kana', 'set-two-lines', 'set-language', 'set-east-asian-language',
      'set-other-language', 'detect-language', 'reset-language-detection',
      'disable-proofing', 'reset-spelling', 'reset-grammar', 'get-proofing', 'show-marks',
    ]],
    ['35-ranges', 'inspect-range-layout-desktop.yaml', [
      'setup', 'get-borders', 'set-shading', 'get-sections', 'get-frames',
      'get-hyperlinks', 'insert-picture', 'get-shapes',
    ]],
    ['35-ranges', 'manage-range-editing-permissions-desktop.yaml', ['setup', 'add-editor', 'get-editors']],
    ['50-document', 'manage-footnotes.yaml', ['setup', 'insert-endnote', 'get-endnotes']],
    ['50-document', 'manage-tracked-changes.yaml', ['get-range-tracked-changes', 'get-range-revisions']],
    ['25-paragraph', 'get-paragraph-on-insertion-point.yaml', ['setup', 'get-paragraph', 'get-sentences']],
    ['05-coauthoring', 'manage-coauthoring-conflicts.yaml', ['get-range-conflicts']],
    ['99-preview-apis', 'read-range-underline.yaml', ['setup', 'get-underline']],
  ] as const;

  test.each(rangeScenarios.flatMap(([group, file, buttons]) =>
    buttons.map(buttonId => ({ group, file, buttonId }))
  ))('Word: $file $buttonId action executes without runtime errors', async ({ group, file, buttonId }) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', group, file),
      buttonId,
      assertions: ({ mockContext }) => {
        expect(consoleErrorSpy).not.toHaveBeenCalled();
        expect(mockContext.sync).toHaveBeenCalled();
      },
    });
  });

  test.each([
    ['05-coauthoring', 'manage-coauthoring-conflicts.yaml', 'get-range-conflicts', 'Conflicts in the selection: 1'],
    ['05-coauthoring', 'manage-coauthoring-conflicts.yaml', 'get-range-conflicts', expect.stringMatching(/^Conflict type: .+$/)],
    ['25-paragraph', 'get-paragraph-on-insertion-point.yaml', 'get-paragraph', expect.stringContaining('The first paragraph of the selected text: ')],
    ['25-paragraph', 'get-paragraph-on-insertion-point.yaml', 'get-sentences', 'Sentences from the current sentence to the paragraph end:'],
    ['35-ranges', 'basic-range-operations.yaml', 'search-range', 'Matches found: 1'],
    ['35-ranges', 'inspect-range-layout-desktop.yaml', 'get-borders', "First paragraph's border formatting:"],
    ['35-ranges', 'inspect-range-layout-desktop.yaml', 'get-sections', 'Sections found: 3'],
    ['35-ranges', 'inspect-range-layout-desktop.yaml', 'get-frames', 'Frames found: 1'],
    ['35-ranges', 'inspect-range-layout-desktop.yaml', 'get-hyperlinks', 'Hyperlinks found: 1'],
    ['35-ranges', 'inspect-range-layout-desktop.yaml', 'get-shapes', 'Shapes found: 1'],
  ] as const)('Word: %s/%s %s logs a single descriptive message', async (group, file, buttonId, message) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', group, file),
      buttonId,
      assertions: () => {
        expect(consoleSpy).toHaveBeenCalledWith(message);
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    ['basic-range-operations.yaml', 'get-range', "First paragraph's properties:", { text: expect.any(String), isEmpty: false, style: 'Normal' }],
    ['format-range-text-desktop.yaml', 'get-formatting', "First paragraph's formatting:", { bold: true, italic: true }],
    ['format-range-text-desktop.yaml', 'get-position', "First paragraph's position and story:", { start: 0, end: 40, storyLength: 80 }],
  ] as const)('Word: %s %s labels structured property output', async (file, buttonId, label, properties) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '35-ranges', file),
      buttonId,
      assertions: () => {
        expect(consoleSpy).toHaveBeenCalledWith(label, expect.objectContaining(properties));
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test('Word: Range.split uses the Range argument order', async () => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '35-ranges', 'split-words-of-first-paragraph.yaml'),
      buttonId: 'highlight',
      assertions: ({ mockRange }) => {
        expect(mockRange.split).toHaveBeenCalledWith([' '], false, true, true);
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test('Word: editing permissions grants access to Everyone on the first paragraph', async () => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '35-ranges', 'manage-range-editing-permissions-desktop.yaml'),
      buttonId: 'add-editor',
      assertions: ({ mockContext, mockBody, mockParagraph, mockRange }) => {
        expect(mockBody.paragraphs.getFirst).toHaveBeenCalled();
        expect(mockParagraph.getRange).toHaveBeenCalledWith('Content');
        expect(mockRange.editors.addByType).toHaveBeenCalledWith('Everyone');
        expect(mockContext.sync).toHaveBeenCalled();
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test('Word: editing permissions reads entries by index and logs loaded properties', async () => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '35-ranges', 'manage-range-editing-permissions-desktop.yaml'),
      buttonId: 'get-editors',
      assertions: ({ mockRange }) => {
        expect(mockRange.editors.getCount).toHaveBeenCalledTimes(1);
        expect(mockRange.editors.getItemAt).toHaveBeenCalledTimes(1);
        expect(mockRange.editors.getItemAt).toHaveBeenCalledWith(0);
        const editor = mockRange.editors.getItemAt.mock.results[0].value;
        expect(editor.load).toHaveBeenCalledWith('name,id');
        expect(consoleSpy).toHaveBeenCalledWith('Editing-permission entries:', 1);
        expect(consoleSpy).toHaveBeenCalledWith({ name: 'Everyone', id: 'Everyone' });
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    ['expand-ranges-safely', 'expandToOrNullObject', 'cannot be combined'],
    ['intersect-ranges-safely', 'intersectWithOrNullObject', 'does not overlap'],
  ] as const)('Word: %s reports a null range', async (buttonId, method, warning) => {
    const { mockObject, mockRange } = createWordMock();
    const nullRange = { ...mockRange, isNullObject: true, load: jest.fn() };
    // Null ranges have no readable text; returning undefined here would conceal an invalid read.
    Object.defineProperty(nullRange, 'text', { get: () => { throw new Error('Null range text is unavailable.'); } });
    mockRange[method].mockReturnValue(nullRange);
    (global as any).Word = mockObject;
    const buttonHandlers = new Map<string, Function>();
    (global as any).document = createMockDocument(buttonHandlers);
    executeSnippetCode(loadSnippetByPath(path.join('samples', 'word', '35-ranges', 'compare-location.yaml')));
    const warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
    await clickButton(buttonHandlers, buttonId);
    expect(warnSpy).toHaveBeenCalledWith(expect.stringContaining(warning));
    expect(consoleErrorSpy).not.toHaveBeenCalled();
  });

  test.each([
    'setup',
    'get-all-tracked-changes',
    'get-first-tracked-change-or-null',
    'get-first-tracked-change-range',
    'get-next-tracked-change',
    'accept-first-tracked-change',
    'reject-first-tracked-change',
    'accept-all-tracked-changes',
    'reject-all-tracked-changes',
  ])('Word: manage tracked changes %s action executes without runtime errors', async (buttonId) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '50-document', 'manage-tracked-changes.yaml'),
      buttonId,
      assertions: ({ mockContext }) => {
        expect(mockContext.sync).toHaveBeenCalled();
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    'setup',
    'get-first-section',
    'get-section-body',
    'get-sections',
    'get-next-section',
    'get-next-section-or-null',
    'add-next-section',
    'add-even-section',
    'add-odd-section',
    'add-continuous-section',
  ])('Word: manage sections %s action executes without runtime errors', async (buttonId) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '50-document', 'insert-section-breaks.yaml'),
      buttonId,
      assertions: ({ mockContext }) => {
        expect(mockContext.sync).toHaveBeenCalled();
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    'setup',
    'apply-layout',
    'apply-page-borders',
    'protect-section',
    'get-page-setup',
    'get-section-protection',
  ])('Word: manage section layout %s action executes without runtime errors', async (buttonId) => {
    (global as any).Office = {
      context: {
        requirements: {
          isSetSupported: jest.fn(() => true),
        },
      },
    };

    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '50-document', 'manage-section-starts.yaml'),
      buttonId,
      assertions: ({ mockContext }) => {
        expect(mockContext.sync).toHaveBeenCalled();
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    'setup',
    'get-row-properties',
    'update-row-data',
    'update-row-formatting',
    'insert-row',
    'delete-row',
    'get-column-properties',
    'update-column-formatting',
    'get-cell-properties',
    'update-cell-data',
    'update-cell-formatting',
  ])('Word: manage table rows, columns, and cells %s action executes without runtime errors', async (buttonId) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '40-tables', 'manage-table-rows-columns-cells.yaml'),
      buttonId,
      assertions: ({ mockContext }) => {
        expect(mockContext.sync).toHaveBeenCalled();
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    {
      snippetPath: path.join('samples', 'word', '42-reference-tables', 'table-of-authorities.yaml'),
      buttonIds: ['setup', 'mark-citations', 'create-table-of-authorities', 'get-properties', 'set-properties', 'delete-table'],
    },
    {
      snippetPath: path.join('samples', 'word', '42-reference-tables', 'table-of-contents.yaml'),
      buttonIds: ['setup', 'create-table-of-contents', 'get-properties', 'set-properties', 'update-page-numbers', 'delete-table'],
    },
    {
      snippetPath: path.join('samples', 'word', '42-reference-tables', 'table-of-figures.yaml'),
      buttonIds: ['setup', 'mark-entries', 'create-table-of-figures', 'get-properties', 'set-properties', 'update-page-numbers', 'delete-table'],
      mockOptions: { paragraphText: 'Figure 1: Quarterly revenue' },
    },
  ])('Word: $snippetPath actions execute without runtime errors', async ({ snippetPath, buttonIds, mockOptions }) => {
    for (const buttonId of buttonIds) {
      await runWordSnippetTest({ snippetPath, buttonId, mockOptions });
    }
  });

  test.each([
    'setup',
    'set-alignment',
    'set-indents',
    'set-spacing',
    'set-pagination',
    'get-paragraph-format',
  ])(
    'Word: paragraph format %s action executes without runtime errors',
    async (buttonId) => {
      await runWordSnippetTest({
        snippetPath: path.join('samples', 'word', '28-formatting', 'paragraph-format.yaml'),
        buttonId,
        assertions: ({ mockContext, mockParagraphFormat, mockStyles }) => {
          expect(mockContext.sync).toHaveBeenCalled();
          expect(
            mockStyles.getByNameOrNullObject.mock.calls.length + mockStyles.getByName.mock.calls.length
          ).toBeGreaterThan(0);

          if (buttonId === 'set-alignment') {
            expect(mockParagraphFormat.alignment).toBe('Justified');
          }

          if (buttonId === 'set-indents') {
            expect(mockParagraphFormat.firstLineIndent).toBe(18);
            expect(mockParagraphFormat.leftIndent).toBe(24);
            expect(mockParagraphFormat.rightIndent).toBe(12);
          }

          if (buttonId === 'set-spacing') {
            expect(mockParagraphFormat.lineSpacing).toBe(18);
            expect(mockParagraphFormat.spaceAfter).toBe(12);
            expect(mockParagraphFormat.spaceBefore).toBe(6);
          }

          if (buttonId === 'set-pagination') {
            expect(mockParagraphFormat.keepTogether).toBe(true);
            expect(mockParagraphFormat.keepWithNext).toBe(true);
            expect(mockParagraphFormat.widowControl).toBe(true);
          }

          if (buttonId === 'get-paragraph-format') {
            expect(mockParagraphFormat.load).toHaveBeenCalled();
            expect(mockParagraphFormat.toJSON).toHaveBeenCalled();
          }

          expect(consoleErrorSpy).not.toHaveBeenCalled();
        },
      });
    }
  );

  test.each([
    'setup',
    'set-border-type',
    'set-border-width',
    'set-border-color',
    'get-border-properties',
    'set-background-color',
    'set-foreground-color',
    'set-texture',
    'get-shading',
  ])(
    'Word: manage styles %s action executes without runtime errors',
    async (buttonId) => {
      await runWordSnippetTest({
        snippetPath: path.join('samples', 'word', '28-formatting', 'manage-styles.yaml'),
        buttonId,
        assertions: ({ mockContext, mockShading, mockStyle, mockStyles }) => {
          expect(mockContext.sync).toHaveBeenCalled();
          expect(
            mockStyles.getByNameOrNullObject.mock.calls.length + mockStyles.getByName.mock.calls.length
          ).toBeGreaterThan(0);

          if (buttonId === 'set-border-type') {
            expect(mockStyle.borders.outsideBorderType).toBe('Dashed');
          }

          if (buttonId === 'set-border-width') {
            expect(mockStyle.borders.outsideBorderWidth).toBe('Pt225');
          }

          if (buttonId === 'set-border-color') {
            expect(mockStyle.borders.outsideBorderColor).toBe('#4472C4');
          }

          if (buttonId === 'get-border-properties') {
            expect(mockStyle.borders.load).toHaveBeenCalledWith(
              'outsideBorderColor, outsideBorderType, outsideBorderWidth'
            );
          }

          if (buttonId === 'set-background-color') {
            expect(mockShading.backgroundPatternColor).toBe('#D9EAF7');
          }

          if (buttonId === 'set-foreground-color') {
            expect(mockShading.foregroundPatternColor).toBe('#1F4E78');
          }

          if (buttonId === 'set-texture') {
            expect(mockShading.texture).toBe('DarkTrellis');
          }

          if (buttonId === 'get-shading') {
            expect(mockShading.load).toHaveBeenCalledWith(
              'backgroundPatternColor, foregroundPatternColor, texture'
            );
          }

          expect(consoleErrorSpy).not.toHaveBeenCalled();
        },
      });
    }
  );
});

describe('Runtime Execution Tests - PowerPoint', () => {
  let consoleSpy: jest.SpyInstance;
  let consoleErrorSpy: jest.SpyInstance;

  beforeEach(() => {
    (global as any).PowerPoint = undefined;
    (global as any).Office = undefined;
    consoleSpy = jest.spyOn(console, 'log').mockImplementation(() => {});
    consoleErrorSpy = jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  test('PowerPoint: basic-api-call-ts executes without runtime errors', async () => {
    await runPowerPointSnippetTest({
      snippetPath: path.join('samples', 'powerpoint', 'basics', 'basic-api-call-ts.yaml'),
      assertions: ({ mockContext, mockShapes }) => {
        // Verify PowerPoint.run was called
        expect(mockContext).toBeDefined();

        // Verify slide operations
        expect(mockContext.presentation.slides.getItemAt).toHaveBeenCalledWith(0);
        expect(mockShapes.addTextBox).toHaveBeenCalledWith('Hello!', expect.any(Object));
        expect(mockContext.sync).toHaveBeenCalled();

        // Verify no errors were logged
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test('PowerPoint: basic-common-api-call executes without runtime errors', async () => {
    await runCommonApiSnippetTest({
      snippetPath: path.join('samples', 'powerpoint', 'basics', 'basic-common-api-call.yaml'),
      assertions: ({ Office }) => {
        // Verify Common API was called
        expect(Office.context.document.getSelectedDataAsync).toHaveBeenCalled();

        // Verify console output
        expect(consoleSpy).toHaveBeenCalled();

        // Verify no errors were logged
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });
});
