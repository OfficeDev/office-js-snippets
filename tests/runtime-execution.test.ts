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
    {
      details: { valueBefore: 'Merchant', valueAfter: 42, valueTypeBefore: 'String', valueTypeAfter: 'Double' },
      expectedLog: 'Change at B2: was Merchant(String), now is 42(Double)',
      isError: false,
    },
    {
      details: { valueBefore: 42, valueAfter: '#DIV/0!', valueTypeBefore: 'Double', valueTypeAfter: 'Error' },
      expectedLog: 'Change at B2: was 42(Double), now is #DIV/0!(Error)',
      isError: true,
    },
    {
      details: undefined,
      expectedLog: 'Edit one table cell at a time to inspect its value and type.',
      isError: false,
    },
  ])('Excel: table change details reports $expectedLog', async ({ details, expectedLog, isError }) => {
    const { mockContext, mockWorksheet } = await runExcelSnippetTest({
      snippetPath: path.join('samples', 'excel', '30-events', 'data-change-event-details.yaml'),
      buttonId: 'register-handler',
    });

    expect(mockContext.workbook.worksheets.getItem).toHaveBeenCalledWith('Sample');
    expect(mockWorksheet.tables.getItem).toHaveBeenCalledWith('ExpensesTable');
    const changedEvent = mockWorksheet.tables.getItem.mock.results[0].value.onChanged;
    expect(changedEvent.add).toHaveBeenCalledTimes(1);
    const handler = changedEvent.add.mock.calls[0][0];
    await handler({ address: 'B2', details });

    expect(consoleSpy).toHaveBeenCalledWith(expectedLog);
    if (isError) {
      expect(consoleSpy).toHaveBeenCalledWith('The updated cell contains an error:', '#DIV/0!');
    } else {
      expect(consoleSpy).not.toHaveBeenCalledWith('The updated cell contains an error:', expect.anything());
    }
    expect(consoleErrorSpy).not.toHaveBeenCalled();
  });
  test.each([
    { basicType: 'Error', basicValue: '#UNKNOWN!' },
    { basicType: 'String', basicValue: 'Unsupported value' },
    { basicType: 'Double', basicValue: 42 },
    { basicType: 'Boolean', basicValue: false },
  ])('Excel: inspects an unavailable value with $basicType fallback', async (fallback) => {
    const unavailableValue = { type: 'NotAvailable', ...fallback };
    await runExcelSnippetTest({
      snippetPath: path.join('samples', 'excel', '20-data-types', 'data-types-error-values.yaml'),
      buttonId: 'inspect-unavailable-value',
      mockOptions: { rangeValuesAsJson: [[unavailableValue]] },
      assertions: ({ mockContext, mockRange, mockWorksheet }) => {
        expect(mockContext.workbook.worksheets.getItem).toHaveBeenCalledWith('Sample');
        expect(mockWorksheet.getRange).toHaveBeenCalledWith('B1');
        expect(mockRange.load).toHaveBeenCalledWith('valuesAsJson');
        expect(mockContext.sync).toHaveBeenCalledTimes(1);
        expect(consoleSpy).toHaveBeenCalledWith('B1 value type:', 'NotAvailable');
        expect(consoleSpy).toHaveBeenCalledWith('B1 fallback type:', fallback.basicType);
        expect(consoleSpy).toHaveBeenCalledWith('B1 fallback value:', fallback.basicValue);
        expect(mockRange.valuesAsJson).toEqual([[unavailableValue]]);
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    { type: 'String', basicValue: 'Sample text' },
    { type: 'LinkedEntity', basicValue: '#VALUE!', text: 'Microsoft Corporation' },
    { type: 'LocalImage', basicValue: '#VALUE!' },
    { type: 'Error', errorType: 'Busy', basicValue: '#BUSY!' },
    { type: 'Error', errorType: 'NotAvailable', basicValue: '#N/A' },
  ])('Excel: does not confuse $basicValue with an unavailable value type', async (cellValue) => {
    await runExcelSnippetTest({
      snippetPath: path.join('samples', 'excel', '20-data-types', 'data-types-error-values.yaml'),
      buttonId: 'inspect-unavailable-value',
      mockOptions: { rangeValuesAsJson: [[cellValue]] },
      assertions: ({ mockRange }) => {
        expect(consoleSpy).toHaveBeenCalledWith('B1 value type:', cellValue.type);
        expect(consoleSpy).toHaveBeenCalledWith(
          "B1's value type is supported. No unavailable-type fallback was returned."
        );
        expect(consoleSpy).toHaveBeenCalledTimes(2);
        expect(consoleSpy).not.toHaveBeenCalledWith('B1 fallback value:', expect.anything());
        expect(mockRange.valuesAsJson).toEqual([[cellValue]]);
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each(['setup', 'set-busy-error', 'change-busy-error'])(
    'Excel: error values %s action executes without runtime errors',
    async (buttonId) => {
      await runExcelSnippetTest({
        snippetPath: path.join('samples', 'excel', '20-data-types', 'data-types-error-values.yaml'),
        buttonId,
        mockOptions: { rangeValuesAsJson: [[{ type: 'Error', errorType: 'Busy' }]] },
        assertions: ({ mockContext, mockRange, mockWorksheet }) => {
          expect(mockContext.sync).toHaveBeenCalled();
          if (buttonId === 'set-busy-error') {
            expect(mockRange.valuesAsJson).toEqual([[{ type: 'Error', errorType: 'Busy' }]]);
          } else if (buttonId === 'change-busy-error') {
            expect(mockWorksheet.getRange).toHaveBeenCalledWith('A1');
            expect(mockRange.values).toEqual([['Process unavailable.']]);
          } else {
            expect(mockRange.values).toEqual([['Ready', 'Sample text']]);
          }
          expect(consoleErrorSpy).not.toHaveBeenCalled();
        },
      });
    }
  );

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

  test.each([
    'setup',
    'add-page-breaks',
    'list-page-breaks',
    'count-page-breaks',
    'get-first-page-break',
    'delete-first-page-break',
    'remove-page-breaks',
  ])('Excel: manage page breaks %s action executes without runtime errors', async (buttonId) => {
    await runExcelSnippetTest({
      snippetPath: path.join('samples', 'excel', '54-worksheet', 'manage-page-breaks.yaml'),
      buttonId,
      assertions: ({ mockContext, mockWorksheet, mockRange }) => {
        const horizontalBreaks = mockWorksheet.horizontalPageBreaks;
        const verticalBreaks = mockWorksheet.verticalPageBreaks;
        expect(mockContext.sync).toHaveBeenCalled();

        switch (buttonId) {
          case 'setup':
            expect(mockContext.workbook.worksheets.add).toHaveBeenCalledWith('Sample');
            expect(mockWorksheet.delete).toHaveBeenCalled();
            expect(mockRange.values).toHaveLength(41);
            expect(mockRange.values.every((row: unknown[]) => row.length === 8)).toBe(true);
            expect(mockRange.values[1]).toEqual(['Product 1', 'South', 110, 110, 110, 110, 440, 2000]);
            expect(mockWorksheet.pageLayout.setPrintArea).toHaveBeenCalledWith('A1:H41');
            expect(mockWorksheet.activate).toHaveBeenCalled();
            break;
          case 'add-page-breaks':
            expect(horizontalBreaks.add.mock.calls).toEqual([['A12'], ['A22']]);
            expect(mockWorksheet.getRange).toHaveBeenCalledWith('E1');
            expect(verticalBreaks.add).toHaveBeenCalledWith(mockRange);
            break;
          case 'list-page-breaks':
            expect(horizontalBreaks.load).toHaveBeenCalledWith('items/rowIndex, items/columnIndex');
            expect(verticalBreaks.load).toHaveBeenCalledWith('items/rowIndex, items/columnIndex');
            expect(consoleSpy).toHaveBeenCalledWith('Horizontal page breaks: 2');
            expect(consoleSpy).toHaveBeenCalledWith('Vertical page breaks: 1');
            expect(consoleSpy).toHaveBeenCalledWith('Horizontal break (zero-based indexes):', {
              rowIndex: 11, columnIndex: 0,
            });
            break;
          case 'count-page-breaks':
            expect(horizontalBreaks.getCount).toHaveBeenCalled();
            expect(verticalBreaks.getCount).toHaveBeenCalled();
            expect(consoleSpy).toHaveBeenCalledWith('Horizontal page breaks: 2');
            expect(consoleSpy).toHaveBeenCalledWith('Vertical page breaks: 1');
            break;
          case 'get-first-page-break':
            expect(horizontalBreaks.getItem).toHaveBeenCalledWith(0);
            expect(horizontalBreaks.items[0].getCellAfterBreak).toHaveBeenCalled();
            expect(horizontalBreaks.items[0].load).toHaveBeenCalledWith('rowIndex, columnIndex');
            expect(mockRange.load).toHaveBeenCalledWith('address');
            expect(consoleSpy).toHaveBeenCalledWith('First horizontal page break (zero-based indexes):', {
              rowIndex: 11, columnIndex: 0, cellAfterBreak: mockRange.address,
            });
            break;
          case 'delete-first-page-break':
            expect(horizontalBreaks.getItem).toHaveBeenCalledWith(0);
            expect(horizontalBreaks.items[0].delete).toHaveBeenCalled();
            expect(verticalBreaks.items[0].delete).not.toHaveBeenCalled();
            break;
          case 'remove-page-breaks':
            expect(horizontalBreaks.removePageBreaks).toHaveBeenCalled();
            expect(verticalBreaks.removePageBreaks).toHaveBeenCalled();
            break;
        }
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test('Excel: function reference setup prepares the Sample worksheet', async () => {
    await runExcelSnippetTest({
      snippetPath: path.join('samples', 'excel', '20-data-types', 'data-types-function-references.yaml'),
      buttonId: 'setup',
      assertions: ({ mockContext, mockRange, mockWorksheet }) => {
        expect(mockContext.workbook.worksheets.getItemOrNullObject).toHaveBeenCalledWith('Sample');
        expect(mockWorksheet.delete).toHaveBeenCalled();
        expect(mockContext.workbook.worksheets.add).toHaveBeenCalledWith('Sample');
        expect(mockRange.values).toEqual([
          ['Entity', 'Purpose'],
          ['Create the reference here', 'Store and inspect a custom-function reference'],
        ]);
        expect(mockWorksheet.activate).toHaveBeenCalled();
        expect(mockContext.sync).toHaveBeenCalled();
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each(['get-first-page-break', 'delete-first-page-break'])(
    'Excel: manage page breaks %s reports an empty collection',
    async (buttonId) => {
      const consoleWarnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
      await runExcelSnippetTest({
        snippetPath: path.join('samples', 'excel', '54-worksheet', 'manage-page-breaks.yaml'),
        buttonId,
        mockOptions: { horizontalPageBreakCount: 0 },
        assertions: ({ mockWorksheet }) => {
          expect(mockWorksheet.horizontalPageBreaks.getItem).not.toHaveBeenCalled();
          expect(consoleWarnSpy).toHaveBeenCalledWith(
            'No horizontal page breaks. Choose Add page breaks first.'
          );
          expect(consoleErrorSpy).not.toHaveBeenCalled();
        },
      });
    }
  );

  test('Excel: page layout add page break action executes without runtime errors', async () => {
    await runExcelSnippetTest({
      snippetPath: path.join('samples', 'excel', '54-worksheet', 'worksheet-page-layout.yaml'),
      buttonId: 'setPageBreaks',
      assertions: ({ mockWorksheet }) => {
        expect(mockWorksheet.horizontalPageBreaks.add).toHaveBeenCalledWith('A21:E21');
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test('Excel: creates a typed custom-function reference inside an entity', async () => {
    await runExcelSnippetTest({
      snippetPath: path.join('samples', 'excel', '20-data-types', 'data-types-function-references.yaml'),
      buttonId: 'create-function-reference',
      assertions: ({ mockContext, mockRange, mockWorksheet }) => {
        expect(mockContext.workbook.worksheets.getItem).toHaveBeenCalledWith('Sample');
        expect(mockWorksheet.getRange).toHaveBeenCalledWith('A2');
        expect(mockRange.valuesAsJson).toEqual([[{
          type: 'Entity',
          text: 'Math value',
          properties: {
            Value: { type: 'Double', basicValue: 10 },
            Get27: {
              type: 'Function',
              functionType: 'JavaScriptReference',
              namespace: 'CONTOSO',
              id: 'GET27',
            },
          },
        }]]);
        expect(mockContext.sync).toHaveBeenCalled();
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test('Excel: reads the stored custom-function identity', async () => {
    await runExcelSnippetTest({
      snippetPath: path.join('samples', 'excel', '20-data-types', 'data-types-function-references.yaml'),
      buttonId: 'get-function-reference',
      mockOptions: {
        rangeValuesAsJson: [[{
          type: 'Entity',
          properties: {
            Get27: {
              type: 'Function',
              functionType: 'JavaScriptReference',
              namespace: 'CONTOSO',
              id: 'GET27',
            },
          },
        }]],
      },
      assertions: ({ mockRange }) => {
        expect(mockRange.load).toHaveBeenCalledWith('valuesAsJson');
        expect(consoleSpy).toHaveBeenCalledWith('Function type:', 'JavaScriptReference');
        expect(consoleSpy).toHaveBeenCalledWith('Custom-function namespace:', 'CONTOSO');
        expect(consoleSpy).toHaveBeenCalledWith('Custom-function ID:', 'GET27');
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each<[unknown, string]>([
    [{ type: 'String', basicValue: 'Create the reference here' }, 'Create the function reference'],
    [{ type: 'Entity', properties: {} }, "doesn't contain a Get27 function reference"],
    [{ type: 'Entity', properties: { Get27: { type: 'Double', basicValue: 27 } } },
      "doesn't contain a Get27 function reference"],
  ])('Excel: reports an invalid function-reference cell (%#)', async (cellValue, message) => {
    await runExcelSnippetTest({
      snippetPath: path.join('samples', 'excel', '20-data-types', 'data-types-function-references.yaml'),
      buttonId: 'get-function-reference',
      mockOptions: { rangeValuesAsJson: [[cellValue]] },
    });

    expect(consoleErrorSpy).toHaveBeenCalledWith(expect.objectContaining({
      message: expect.stringContaining(message),
    }));
    expect(consoleSpy).not.toHaveBeenCalled();
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

  test.each([
    'setup',
    'get-list-string',
    'get-sibling-index',
    'get-ancestor',
    'get-ancestor-or-null',
    'get-descendants',
    'set-level',
  ])('Word: manage list items %s action executes without runtime errors', async (buttonId) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '20-lists', 'manage-list-items.yaml'),
      buttonId,
      assertions: ({ mockContext, mockBody, mockParagraph }) => {
        const list = mockBody.lists.getFirst();
        const listItem = mockParagraph.listItem;
        if (buttonId === 'setup') {
          expect(mockBody.clear).toHaveBeenCalled();
          expect(mockBody.insertParagraph).toHaveBeenCalledWith('Design', 'Start');
          expect(mockParagraph.startNewList).toHaveBeenCalled();
          expect(list.insertParagraph.mock.calls).toEqual([
            ['Draft', 'End'], ['Review', 'End'], ['Publish', 'End'],
          ]);
          expect(list.setLevelNumbering.mock.calls).toEqual([
            [0, 'Arabic'], [1, 'LowerLetter'], [2, 'LowerRoman'],
          ]);
        } else if (buttonId === 'get-list-string') {
          expect(list.paragraphs.getLast).toHaveBeenCalled();
          expect(listItem.load).toHaveBeenCalledWith('listString');
          expect(consoleSpy).toHaveBeenCalledWith('Last list item marker: b.');
        } else if (buttonId === 'get-sibling-index') {
          expect(list.paragraphs.getLast).toHaveBeenCalled();
          expect(listItem.load).toHaveBeenCalledWith('siblingIndex');
          expect(consoleSpy).toHaveBeenCalledWith('Last list item sibling index: 1');
        } else if (buttonId === 'get-ancestor') {
          expect(listItem.getAncestor).toHaveBeenCalledWith(true);
          expect(mockParagraph.load).toHaveBeenCalledWith('text');
          expect(consoleSpy).toHaveBeenCalledWith('Parent of the last list item: Mock paragraph');
        } else if (buttonId === 'get-ancestor-or-null') {
          expect(listItem.getAncestorOrNullObject).toHaveBeenCalledWith();
          expect(consoleSpy).toHaveBeenCalledWith(
            'The first list item has no ancestor because it is at the top level.'
          );
        } else if (buttonId === 'get-descendants') {
          expect(listItem.getDescendants.mock.calls).toEqual([[true], [false]]);
          expect(list.paragraphs.load).toHaveBeenCalledWith('items/text');
          expect(consoleSpy).toHaveBeenCalledWith('Direct children (2):');
          expect(consoleSpy).toHaveBeenCalledWith('All descendants (2):');
          expect(consoleSpy).toHaveBeenCalledWith('- Mock paragraph');
        } else {
          expect(listItem.level).toBe(2);
        }
        expect(mockContext.sync).toHaveBeenCalled();
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([false, true])('Word: optional list ancestor handles isNullObject=%s', async (isNullObject) => {
    const { mockObject, mockContext, mockParagraph } = createWordMock();
    const ancestor = { isNullObject, text: 'Design', load: jest.fn() };
    mockParagraph.listItem.getAncestorOrNullObject.mockReturnValue(ancestor);
    (global as any).Word = mockObject;
    const buttonHandlers = new Map<string, Function>();
    (global as any).document = createMockDocument(buttonHandlers);
    executeSnippetCode(loadSnippetByPath(path.join('samples', 'word', '20-lists', 'manage-list-items.yaml')));

    await clickButton(buttonHandlers, 'get-ancestor-or-null');

    expect(ancestor.load).toHaveBeenCalledWith('text');
    expect(mockContext.sync).toHaveBeenCalled();
    expect(consoleSpy).toHaveBeenCalledWith(isNullObject
      ? 'The first list item has no ancestor because it is at the top level.'
      : 'Ancestor of the first list item: Design');
    expect(consoleErrorSpy).not.toHaveBeenCalled();
  });

  test('Word: get document breaks executes without runtime errors', async () => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '36-pages', 'get-pages.yaml'),
      buttonId: 'get-breaks',
      assertions: ({ mockContext }) => {
        const breaks = mockContext.document.activeWindow.activePane.pages.items[0].breaks;
        expect(breaks.load).toHaveBeenCalledWith('items/pageIndex');
        expect(consoleSpy).toHaveBeenCalledWith('Breaks found: 1');
        expect(consoleSpy).toHaveBeenCalledWith('Break on page 1.');
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    { label: 'a page break character', rangeText: '\f', escapedText: '"\\f"' },
    { label: 'an empty range', rangeText: '', escapedText: '""' },
    { label: 'text with a newline', rangeText: 'First line\nSecond line', escapedText: '"First line\\nSecond line"' },
  ])('Word: get break ranges displays $label', async ({ rangeText, escapedText }) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '36-pages', 'get-pages.yaml'),
      buttonId: 'get-break-ranges',
      mockOptions: { selectionText: rangeText },
      assertions: ({ mockContext, mockRange }) => {
        const pages = mockContext.document.activeWindow.activePane.pages;
        expect(pages.load).toHaveBeenCalledWith('items');
        expect(pages.items[0].breaks.load).toHaveBeenCalledWith('items');
        expect(mockRange.load).toHaveBeenCalledWith('text');
        expect(consoleSpy.mock.calls).toEqual([
          ['Break ranges found: 1'],
          [`Break 1 range text (escaped): ${escapedText}`],
        ]);
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    'setup',
    'get-pages-selected-range',
    'get-pages-enclosing-viewport',
    'get-all-pages',
    'get-page-dimensions',
    'get-next-page',
    'get-page-data',
    'get-all-panes',
    'get-all-windows',
  ])('Word: pages %s action executes without runtime errors', async (buttonId) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '36-pages', 'get-pages.yaml'),
      buttonId,
      assertions: ({ mockContext, mockBody, mockRange }) => {
        const pages = mockContext.document.activeWindow.activePane.pages;
        if (buttonId === 'setup') {
          expect(mockBody.insertText).toHaveBeenCalledWith(
            'Page one: Introduction to the page sample.', 'Replace'
          );
          expect(mockBody.insertBreak).toHaveBeenCalledTimes(2);
          expect(mockRange.select).toHaveBeenCalled();
        } else if (buttonId === 'get-page-dimensions') {
          expect(pages.items[0].load).toHaveBeenCalledWith('index,height,width');
          expect(consoleSpy).toHaveBeenCalledWith('First page properties:', {
            index: 1, heightInPoints: 792, widthInPoints: 612,
          });
        } else if (buttonId === 'get-next-page') {
          expect(pages.items[0].getNext).toHaveBeenCalled();
          expect(consoleSpy).toHaveBeenCalledWith('Next page index: 2');
        } else if (buttonId === 'get-page-data') {
          expect(pages.items[0].toJSON).toHaveBeenCalled();
          expect(consoleSpy).toHaveBeenCalledWith('First page data:', {
            index: 1, height: 792, width: 612,
          });
        }
        expect(mockContext.sync).toHaveBeenCalled();
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([true, false])('Word: third-paragraph pages handles paragraph availability (%s)', async (hasThirdParagraph) => {
    const { mockObject, mockParagraphs, mockParagraph } = createWordMock();
    if (hasThirdParagraph) {
      mockParagraphs.items.push({ ...mockParagraph });
    }
    const warningSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
    (global as any).Word = mockObject;
    const buttonHandlers = new Map<string, Function>();
    (global as any).document = createMockDocument(buttonHandlers);
    executeSnippetCode(loadSnippetByPath(path.join('samples', 'word', '36-pages', 'get-pages.yaml')));

    await clickButton(buttonHandlers, 'get-pages-third-paragraph');

    if (hasThirdParagraph) {
      expect(consoleSpy).toHaveBeenCalledWith('Pages containing the third paragraph: 2');
      expect(warningSpy).not.toHaveBeenCalled();
    } else {
      expect(warningSpy).toHaveBeenCalledWith('Run Set up to create at least three paragraphs.');
    }
    expect(consoleErrorSpy).not.toHaveBeenCalled();
  });

  test('Word: next page warns instead of calling getNext on a single-page document', async () => {
    const { mockObject, mockPages, mockPage } = createWordMock();
    mockPages.items.splice(1);
    const warningSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
    (global as any).Word = mockObject;
    const buttonHandlers = new Map<string, Function>();
    (global as any).document = createMockDocument(buttonHandlers);
    executeSnippetCode(loadSnippetByPath(path.join('samples', 'word', '36-pages', 'get-pages.yaml')));

    await clickButton(buttonHandlers, 'get-next-page');

    expect(mockPage.getNext).not.toHaveBeenCalled();
    expect(warningSpy).toHaveBeenCalledWith('Run Set up to create multiple pages before getting the next page.');
    expect(consoleErrorSpy).not.toHaveBeenCalled();
  });

  test.each(['setup', 'iterate-pages'])('Word: page iteration %s action executes without runtime errors', async (buttonId) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '36-pages', 'iterate-pages.yaml'),
      buttonId,
      assertions: ({ mockContext, mockBody }) => {
        const pages = mockContext.document.activeWindow.activePane.pages;
        if (buttonId === 'setup') {
          expect(mockBody.insertText).toHaveBeenCalledWith(
            'Page one: Introduction to page iteration.', 'Replace'
          );
          expect(mockBody.insertBreak).toHaveBeenCalledTimes(2);
        } else {
          expect(pages.getFirstOrNullObject).toHaveBeenCalled();
          expect(pages.items[0].getNextOrNullObject).toHaveBeenCalledTimes(1);
          expect(pages.items[1].getNextOrNullObject).toHaveBeenCalledTimes(1);
          expect(consoleSpy).toHaveBeenCalledWith('Reached the end of the pages in the active pane.');
        }
        expect(mockContext.sync).toHaveBeenCalled();
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  const rangeScenarios = [
    ['25-paragraph', 'basic-paragraph-operations.yaml', [
      'setup', 'get-paragraph', 'get-next', 'get-previous', 'check-next', 'check-previous',
      'get-sentences', 'clear-paragraph', 'delete-paragraph',
    ]],
    ['25-paragraph', 'insert-and-read-paragraph-markup.yaml', [
      'setup', 'insert-html', 'insert-ooxml', 'get-html', 'get-ooxml',
    ]],
    ['25-paragraph', 'paragraph-properties.yaml', [
      'setup', 'indent', 'indent-first-line', 'indent-right', 'spacing', 'space-before',
      'space-after', 'line-unit-spacing-after', 'line-unit-spacing-before', 'align', 'get-formatting',
    ]],
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
      'set-bookmark-start', 'set-bookmark-end', 'copy-bookmark', 'select-bookmark',
      'check-bookmark',
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
    ['25-paragraph', 'basic-paragraph-operations.yaml', 'get-paragraph', expect.stringContaining('First paragraph: ')],
    ['25-paragraph', 'basic-paragraph-operations.yaml', 'check-next', 'There is no paragraph after the last paragraph.'],
    ['25-paragraph', 'basic-paragraph-operations.yaml', 'check-previous', 'There is no paragraph before the first paragraph.'],
    ['25-paragraph', 'basic-paragraph-operations.yaml', 'get-next', expect.stringContaining('Next paragraph: ')],
    ['25-paragraph', 'basic-paragraph-operations.yaml', 'get-previous', expect.stringContaining('Previous paragraph: ')],
    ['25-paragraph', 'basic-paragraph-operations.yaml', 'get-sentences', 'Sentences found: 1'],
    ['25-paragraph', 'insert-and-read-paragraph-markup.yaml', 'get-html', 'Paragraph HTML: <p>Mock paragraph</p>'],
    ['25-paragraph', 'insert-and-read-paragraph-markup.yaml', 'get-ooxml', 'Paragraph OOXML: <w:p>Mock paragraph</w:p>'],
    ['05-coauthoring', 'manage-coauthors.yaml', 'get-coauthoring-status', 'Can coauthor: true'],
    ['05-coauthoring', 'manage-coauthors.yaml', 'get-coauthoring-status', 'Can merge: true'],
    ['05-coauthoring', 'manage-coauthors.yaml', 'get-coauthoring-status', 'Pending updates: false'],
    ['05-coauthoring', 'manage-coauthors.yaml', 'get-my-info', expect.stringMatching(/^Current user: .+$/)],
    ['05-coauthoring', 'manage-coauthoring-conflicts.yaml', 'get-range-conflicts', 'Conflicts in the selection: 1'],
    ['05-coauthoring', 'manage-coauthoring-conflicts.yaml', 'get-range-conflicts', expect.stringMatching(/^Conflict type: .+$/)],
    ['25-paragraph', 'get-paragraph-on-insertion-point.yaml', 'get-paragraph', expect.stringContaining('The first paragraph of the selected text: ')],
    ['25-paragraph', 'get-paragraph-on-insertion-point.yaml', 'get-sentences', 'Sentences from the current sentence to the paragraph end:'],
    ['35-ranges', 'basic-range-operations.yaml', 'search-range', 'Matches found: 1'],
    ['35-ranges', 'compare-location.yaml', 'expand-ranges', expect.stringContaining('Expanded text: ')],
    ['35-ranges', 'compare-location.yaml', 'intersect-ranges', expect.stringContaining('Overlapping text: ')],
    ['35-ranges', 'format-range-text-desktop.yaml', 'detect-language', 'Language detected: true'],
    ['35-ranges', 'insert-and-read-formatted-content.yaml', 'get-html', expect.stringContaining('Range HTML: ')],
    ['35-ranges', 'insert-and-read-formatted-content.yaml', 'get-ooxml', expect.stringContaining('Range OOXML: ')],
    ['35-ranges', 'inspect-range-layout-desktop.yaml', 'get-borders', "First paragraph's border formatting:"],
    ['35-ranges', 'inspect-range-layout-desktop.yaml', 'get-sections', 'Sections found: 3'],
    ['35-ranges', 'inspect-range-layout-desktop.yaml', 'get-frames', 'Frames found: 1'],
    ['35-ranges', 'inspect-range-layout-desktop.yaml', 'get-hyperlinks', 'Hyperlinks found: 1'],
    ['35-ranges', 'inspect-range-layout-desktop.yaml', 'get-shapes', 'Shapes found: 1'],
    ['35-ranges', 'inspect-range-structure.yaml', 'get-controls', 'Content controls found: 1'],
    ['35-ranges', 'inspect-range-structure.yaml', 'get-fields', 'Fields found: 1'],
    ['35-ranges', 'inspect-range-structure.yaml', 'get-pictures', 'Inline pictures found: 1'],
    ['35-ranges', 'inspect-range-structure.yaml', 'get-lists', 'Lists found: 1'],
    ['35-ranges', 'inspect-range-structure.yaml', 'get-lists', 'List ID: 1'],
    ['35-ranges', 'inspect-range-structure.yaml', 'get-tables', 'Tables found: 1'],
    ['35-ranges', 'inspect-range-structure.yaml', 'get-hyperlink-ranges', 'Hyperlinks found: 1'],
    ['35-ranges', 'manage-range-bookmarks.yaml', 'get-bookmarks', 'Bookmarks found: 1'],
    ['35-ranges', 'split-words-of-first-paragraph.yaml', 'get-next-sentence', expect.stringContaining('Next sentence: ')],
    ['50-document', 'manage-footnotes.yaml', 'get-endnotes', 'Endnotes found: 1'],
    ['50-document', 'manage-tracked-changes.yaml', 'get-range-revisions', 'Revisions in the first paragraph: 1'],
    ['50-document', 'manage-tracked-changes.yaml', 'get-range-tracked-changes', expect.stringMatching(/^Tracked changes in the first paragraph: \d+$/)],
    ['50-document', 'manage-tracked-changes.yaml', 'get-first-tracked-change-range', expect.stringContaining('Tracked change range text: ')],
    ['99-preview-apis', 'read-range-underline.yaml', 'get-underline', 'Underline type: Single'],
    ['99-preview-apis', 'read-range-underline.yaml', 'get-underline', 'Has a single underline: true'],
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
    ['25-paragraph', 'paragraph-properties.yaml', 'get-formatting', 'Paragraph:', { firstLineIndent: 0, leftIndent: 0, rightIndent: 0, spaceBefore: 0 }],
    ['35-ranges', 'basic-range-operations.yaml', 'get-range', "First paragraph's properties:", { text: expect.any(String), isEmpty: false, style: 'Normal' }],
    ['35-ranges', 'format-range-text-desktop.yaml', 'get-formatting', "First paragraph's formatting:", { bold: true, italic: true }],
    ['35-ranges', 'format-range-text-desktop.yaml', 'get-position', "First paragraph's position and story:", { start: 0, end: 40, storyLength: 80 }],
    ['35-ranges', 'format-range-text-desktop.yaml', 'get-proofing', "First paragraph's proofing settings:", { languageId: 'EnglishUS', languageDetected: false }],
    ['35-ranges', 'inspect-range-structure.yaml', 'get-parent-table', 'Parent table:', { rowCount: 2, values: expect.any(Array) }],
    ['35-ranges', 'inspect-range-structure.yaml', 'get-parent-cell', 'Parent cell:', { rowIndex: expect.any(Number), cellIndex: expect.any(Number), value: expect.any(String) }],
    ['50-document', 'manage-tracked-changes.yaml', 'get-next-tracked-change', 'Next tracked change:', { author: expect.any(String), date: expect.any(Date), text: expect.any(String), type: expect.any(String) }],
  ] as const)('Word: %s/%s %s labels structured property output', async (group, file, buttonId, label, properties) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', group, file),
      buttonId,
      assertions: () => {
        expect(consoleSpy).toHaveBeenCalledWith(label, expect.objectContaining(properties));
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    'get-bookmarks', 'set-bookmark-start', 'set-bookmark-end',
    'copy-bookmark', 'select-bookmark', 'check-bookmark', 'delete-bookmark',
  ])('Word: bookmark %s uses the typed bookmark operation', async (buttonId) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '35-ranges', 'manage-range-bookmarks.yaml'),
      buttonId,
      mockOptions: {
        bookmarkNames: buttonId === 'delete-bookmark'
          ? ['RangeSampleSummary', 'RangeSampleCopy'] : ['RangeSampleSummary'],
      },
      assertions: ({ mockContext, mockRange }) => {
        const bookmarks = mockRange.bookmarks;
        const summaryBookmark = bookmarks.items[0];
        expect(mockContext.sync).toHaveBeenCalled();
        switch (buttonId) {
          case 'get-bookmarks':
            expect(bookmarks.load).toHaveBeenCalledWith(
              'items/name,items/start,items/end,items/isEmpty,items/isColumn,items/storyType,items/range/text'
            );
            expect(consoleSpy).toHaveBeenCalledWith('Bookmark details:', {
              name: 'RangeSampleSummary', start: 0, end: 40,
              isEmpty: false, isColumn: false, storyType: 'MainText', text: mockRange.text,
            });
            break;
          case 'set-bookmark-start':
            expect(mockRange.load).toHaveBeenCalledWith('start');
            expect(bookmarks.load).toHaveBeenCalledWith('items/name,items/end');
            expect(summaryBookmark.start).toBe(8);
            expect(summaryBookmark.end).toBe(40);
            break;
          case 'set-bookmark-end':
            expect(mockRange.load).toHaveBeenCalledWith('end');
            expect(bookmarks.load).toHaveBeenCalledWith('items/name,items/start');
            expect(summaryBookmark.end).toBe(39);
            expect(summaryBookmark.start).toBe(0);
            break;
          case 'copy-bookmark':
            expect(summaryBookmark.copyTo).toHaveBeenCalledWith('RangeSampleCopy');
            break;
          case 'select-bookmark':
            expect(summaryBookmark.select).toHaveBeenCalled();
            break;
          case 'check-bookmark':
            expect(bookmarks.exists).toHaveBeenCalledWith('RangeSampleCopy');
            expect(consoleSpy).toHaveBeenCalledWith('RangeSampleCopy exists: false');
            break;
          case 'delete-bookmark':
            expect(bookmarks.items[1].delete).toHaveBeenCalled();
            expect(summaryBookmark.delete).not.toHaveBeenCalled();
            expect(mockRange.delete).not.toHaveBeenCalled();
            break;
        }
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    ['clear-paragraph', 'clear', 'delete'],
    ['delete-paragraph', 'delete', 'clear'],
  ])('Word: %s invokes the paragraph operation, not a range operation', async (buttonId, method, otherMethod) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '25-paragraph', 'basic-paragraph-operations.yaml'),
      buttonId,
      assertions: ({ mockParagraph, mockRange }) => {
        expect(mockParagraph[method]).toHaveBeenCalledTimes(1);
        expect(mockParagraph[otherMethod]).not.toHaveBeenCalled();
        expect(mockRange[method]).not.toHaveBeenCalled();
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    'set-bookmark-start', 'set-bookmark-end', 'copy-bookmark', 'select-bookmark', 'delete-bookmark',
  ])('Word: bookmark %s reports a missing bookmark', async (buttonId) => {
    const warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '35-ranges', 'manage-range-bookmarks.yaml'),
      buttonId,
      mockOptions: { bookmarkNames: [] },
      assertions: () => {
        expect(warnSpy).toHaveBeenCalledWith(
          buttonId === 'delete-bookmark'
            ? 'RangeSampleCopy was not found. Choose Copy bookmark first.'
            : 'RangeSampleSummary was not found. Choose Insert bookmark first.'
        );
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    ['indent-first-line', 'firstLineIndent', 36],
    ['indent-right', 'rightIndent', 72],
    ['space-before', 'spaceBefore', 24],
  ])('Word: %s assigns the requested paragraph property', async (buttonId, property, value) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '25-paragraph', 'paragraph-properties.yaml'),
      buttonId: String(buttonId),
      assertions: ({ mockParagraph }) => {
        expect(mockParagraph[property]).toBe(value);
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    ['insert-html', 'insertHtml', expect.stringContaining('<b>Project update:</b>')],
    ['insert-ooxml', 'insertOoxml', expect.stringContaining('<pkg:package')],
  ])('Word: %s replaces content through the Paragraph API', async (buttonId, method, markup) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '25-paragraph', 'insert-and-read-paragraph-markup.yaml'),
      buttonId,
      assertions: ({ mockParagraph, mockRange }) => {
        expect(mockParagraph[method]).toHaveBeenCalledWith(markup, 'Replace');
        expect(mockRange[method]).not.toHaveBeenCalled();
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test('Word: bookmark copy reports an existing copy without overwriting it', async () => {
    const warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '35-ranges', 'manage-range-bookmarks.yaml'),
      buttonId: 'copy-bookmark',
      mockOptions: { bookmarkNames: ['RangeSampleSummary', 'RangeSampleCopy'] },
      assertions: ({ mockRange }) => {
        expect(mockRange.bookmarks.items[0].copyTo).not.toHaveBeenCalled();
        expect(warnSpy).toHaveBeenCalledWith(
          'RangeSampleCopy already exists. Delete the copied bookmark before copying again.'
        );
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    ['check-next', 'getNextOrNullObject', 'Next paragraph: Neighbor paragraph.'],
    ['check-previous', 'getPreviousOrNullObject', 'Previous paragraph: Neighbor paragraph.'],
  ])('Word: %s handles both null and non-null paragraphs', async (buttonId, method, message) => {
    const { mockObject, mockParagraph } = createWordMock();
    const neighbor = { isNullObject: false, text: 'Neighbor paragraph.', load: jest.fn() };
    const getNeighbor = jest.fn(() => neighbor);
    Object.assign(mockParagraph, { [method]: getNeighbor });
    (global as any).Word = mockObject;
    const buttonHandlers = new Map<string, Function>();
    (global as any).document = createMockDocument(buttonHandlers);
    executeSnippetCode(loadSnippetByPath(path.join('samples', 'word', '25-paragraph', 'basic-paragraph-operations.yaml')));
    await clickButton(buttonHandlers, buttonId);
    expect(consoleSpy).toHaveBeenCalledWith(message);

    neighbor.isNullObject = true;
    Object.defineProperty(neighbor, 'text', { get: () => { throw new Error('Null paragraph text is unavailable.'); } });
    await clickButton(buttonHandlers, buttonId);
    expect(consoleErrorSpy).not.toHaveBeenCalled();
    expect(getNeighbor).toHaveBeenCalledTimes(2);
  });

  test('Word: paragraph sentences use the paragraph-specific argument order', async () => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '25-paragraph', 'basic-paragraph-operations.yaml'),
      buttonId: 'get-sentences',
      assertions: ({ mockParagraph }) => {
        expect(mockParagraph.getTextRanges).toHaveBeenCalledWith(['.'], true);
        const sentences = mockParagraph.getTextRanges.mock.results[0].value;
        expect(sentences.load).toHaveBeenCalledWith('items/text');
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test('Word: bookmark existence reports a present copy', async () => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '35-ranges', 'manage-range-bookmarks.yaml'),
      buttonId: 'check-bookmark',
      mockOptions: { bookmarkNames: ['RangeSampleSummary', 'RangeSampleCopy'] },
      assertions: () => {
        expect(consoleSpy).toHaveBeenCalledWith('RangeSampleCopy exists: true');
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each(['set-bookmark-start', 'set-bookmark-end'])(
    'Word: bookmark %s rejects a boundary that would cross the other endpoint',
    async (buttonId) => {
      const { mockObject, mockRange, mockContext } = createWordMock();
      const summaryBookmark = mockRange.bookmarks.items[0];
      if (buttonId === 'set-bookmark-start') {
        summaryBookmark.end = 8;
      } else {
        summaryBookmark.start = 39;
      }
      (global as any).Word = mockObject;
      const buttonHandlers = new Map<string, Function>();
      (global as any).document = createMockDocument(buttonHandlers);
      executeSnippetCode(loadSnippetByPath(
        path.join('samples', 'word', '35-ranges', 'manage-range-bookmarks.yaml')
      ));
      const warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
      await clickButton(buttonHandlers, buttonId);
      expect(warnSpy).toHaveBeenCalledWith(
        'The bookmark is too short. Run Set up and Insert bookmark again.'
      );
      expect(summaryBookmark.start).toBe(buttonId === 'set-bookmark-start' ? 0 : 39);
      expect(summaryBookmark.end).toBe(buttonId === 'set-bookmark-start' ? 8 : 40);
      expect(mockContext.sync).toHaveBeenCalledTimes(1);
      expect(consoleErrorSpy).not.toHaveBeenCalled();
    }
  );

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
        expect(consoleSpy).toHaveBeenCalledWith('Editing-permission entries: 1');
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

  test('Word: revision metadata includes formatting and collection position', async () => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '50-document', 'manage-tracked-changes.yaml'),
      buttonId: 'get-range-revisions',
      assertions: ({ mockRange }) => {
        expect(mockRange.revisions.load).toHaveBeenCalledWith(
          'items/author,items/date,items/type,items/index,items/formatDescription'
        );
        expect(consoleSpy).toHaveBeenCalledWith('Revision properties:', {
          author: 'Mock Author',
          date: new Date('2024-01-01'),
          type: 'Insert',
          index: 1,
          formatDescription: '',
        });
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    ['get-first-revision-range', true],
    ['get-first-revision-range', false],
    ['accept-first-revision', true],
    ['accept-first-revision', false],
    ['reject-first-revision', true],
    ['reject-first-revision', false],
  ] as const)('Word: %s handles revision availability (%s)', async (buttonId, hasRevision) => {
    const { mockObject, mockContext, mockRange } = createWordMock();
    const revision = mockRange.revisions.items[0];
    if (!hasRevision) {
      mockRange.revisions.items.length = 0;
    }
    const warningSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
    (global as any).Word = mockObject;
    const buttonHandlers = new Map<string, Function>();
    (global as any).document = createMockDocument(buttonHandlers);
    executeSnippetCode(loadSnippetByPath(path.join('samples', 'word', '50-document', 'manage-tracked-changes.yaml')));

    await clickButton(buttonHandlers, buttonId);

    expect(mockRange.revisions.load).toHaveBeenCalledWith('items');
    if (!hasRevision) {
      expect(warningSpy).toHaveBeenCalledWith('No revisions were found. Run Set up to create tracked changes.');
      expect(mockRange.select).not.toHaveBeenCalled();
      expect(revision.accept).not.toHaveBeenCalled();
      expect(revision.reject).not.toHaveBeenCalled();
    } else if (buttonId === 'get-first-revision-range') {
      expect(mockRange.load).toHaveBeenCalledWith('text');
      expect(mockRange.select).toHaveBeenCalled();
      expect(consoleSpy).toHaveBeenCalledWith(`First revision range text: ${mockRange.text}`);
    } else if (buttonId === 'accept-first-revision') {
      expect(revision.accept).toHaveBeenCalledTimes(1);
      expect(revision.reject).not.toHaveBeenCalled();
    } else {
      expect(revision.reject).toHaveBeenCalledTimes(1);
      expect(revision.accept).not.toHaveBeenCalled();
    }
    expect(mockContext.sync).toHaveBeenCalled();
    expect(consoleErrorSpy).not.toHaveBeenCalled();
  });

  test('Word: tracked move setup creates two paragraphs and enables move tracking', async () => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '50-document', 'track-moved-text.yaml'),
      buttonId: 'setup',
      assertions: ({ mockContext, mockBody, mockRange }) => {
        expect(mockBody.clear).toHaveBeenCalled();
        expect(mockBody.insertParagraph).toHaveBeenNthCalledWith(
          1, 'This paragraph stays in place.', 'Start'
        );
        expect(mockBody.insertParagraph).toHaveBeenNthCalledWith(
          2, 'Move this entire paragraph above the first paragraph.', 'End'
        );
        expect(mockContext.document.trackMoves).toBe(true);
        expect(mockContext.document.changeTrackingMode).toBe('TrackAll');
        expect(mockRange.select).toHaveBeenCalled();
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each(['MovedFrom', 'MovedTo', 'Insert', 'Empty'])(
    'Word: moved revision range handles %s revisions',
    async (revisionType) => {
      const { mockObject, mockContext, mockRange } = createWordMock();
      const revisions = mockContext.document.revisions;
      const revision = revisions.items[0];
      mockRange.revisions.items.length = 0;
      revision.type = revisionType;
      if (revisionType === 'Empty') {
        revisions.items.length = 0;
      }
      const movedRange = { ...mockRange, text: 'Moved sample text', load: jest.fn(), select: jest.fn() };
      const movedRangeGetter = jest.spyOn(revision, 'movedRange', 'get').mockReturnValue(movedRange);
      const warningSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
      (global as any).Word = mockObject;
      const buttonHandlers = new Map<string, Function>();
      (global as any).document = createMockDocument(buttonHandlers);
      executeSnippetCode(loadSnippetByPath(path.join('samples', 'word', '50-document', 'track-moved-text.yaml')));

      await clickButton(buttonHandlers, 'get-moved-revision-range');

      expect(revisions.load).toHaveBeenCalledWith('items/type');
      expect(mockRange.revisions.load).not.toHaveBeenCalled();
      if (revisionType === 'MovedFrom' || revisionType === 'MovedTo') {
        expect(movedRangeGetter).toHaveBeenCalled();
        expect(movedRange.load).toHaveBeenCalledWith('text');
        expect(movedRange.select).toHaveBeenCalled();
        expect(consoleSpy).toHaveBeenCalledWith('Moved revision range text: Moved sample text');
        expect(warningSpy).not.toHaveBeenCalled();
      } else {
        expect(movedRangeGetter).not.toHaveBeenCalled();
        expect(movedRange.select).not.toHaveBeenCalled();
        expect(consoleSpy).toHaveBeenCalledWith(
          'Revision types found:', revisionType === 'Empty' ? [] : ['Insert']
        );
        expect(warningSpy).toHaveBeenCalledWith(
          'No tracked move was found. Move an entire paragraph as described in the task pane; moving a phrase may produce insertion and deletion revisions instead.'
        );
      }
      expect(consoleErrorSpy).not.toHaveBeenCalled();
    }
  );

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
