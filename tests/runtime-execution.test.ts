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

  test('Word: page sample setup creates a repeatable document with a page break', async () => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '35-ranges', 'get-pages.yaml'),
      buttonId: 'setup',
      assertions: ({ mockContext, mockBody }) => {
        expect(mockBody.clear).toHaveBeenCalled();
        expect(mockBody.insertBreak).toHaveBeenCalledWith('Page', 'End');
        expect(mockBody.insertParagraph).toHaveBeenCalled();
        expect(mockContext.sync).toHaveBeenCalled();
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    { label: 'multiple pages and breaks', pageBreakIndexes: [[1], [], [3, 3]] },
    { label: 'pages without breaks', pageBreakIndexes: [[], []] },
    { label: 'an empty page collection', pageBreakIndexes: [] },
  ])('Word: get breaks handles $label', async ({ pageBreakIndexes }) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '35-ranges', 'get-pages.yaml'),
      buttonId: 'get-breaks',
      mockOptions: { pageBreakIndexes },
      assertions: ({ mockContext }) => {
        const pages = mockContext.document.activeWindow.activePane.pages;
        expect(pages.load).toHaveBeenCalledWith('items');
        for (const page of pages.items) {
          expect(page.breaks.load).toHaveBeenCalledWith('items/pageIndex');
        }
        const breakIndexes = pageBreakIndexes.flat();
        expect(consoleSpy.mock.calls).toEqual([
          [`Breaks found: ${breakIndexes.length}`],
          ...breakIndexes.map(pageIndex => [`Break on page ${pageIndex}.`]),
        ]);
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    { label: 'multiple pages and breaks', pageBreakIndexes: [[1], [], [3, 3]] },
    { label: 'pages without breaks', pageBreakIndexes: [[], []] },
    { label: 'an empty page collection', pageBreakIndexes: [] },
  ])('Word: get break ranges handles $label', async ({ pageBreakIndexes }) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '35-ranges', 'get-pages.yaml'),
      buttonId: 'get-break-ranges',
      mockOptions: { pageBreakIndexes },
      assertions: ({ mockContext }) => {
        const pages = mockContext.document.activeWindow.activePane.pages;
        expect(pages.load).toHaveBeenCalledWith('items');
        for (const page of pages.items) {
          expect(page.breaks.load).toHaveBeenCalledWith('items');
          for (const documentBreak of page.breaks.items) {
            expect(documentBreak.range.load).toHaveBeenCalledWith('text');
          }
        }
        const breakCount = pageBreakIndexes.flat().length;
        expect(consoleSpy.mock.calls).toEqual([
          [`Break ranges found: ${breakCount}`],
          ...Array.from({ length: breakCount }, (_, index) => [
            `Break ${index + 1} range text (escaped): "\\f"`,
          ]),
        ]);
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
