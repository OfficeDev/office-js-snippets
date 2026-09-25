import {
  runExcelSnippetTest,
  runWordSnippetTest,
  runPowerPointSnippetTest,
  runCommonApiSnippetTest,
} from './helpers/snippet-test-runner';
import * as path from 'path';

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
