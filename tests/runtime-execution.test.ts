import {
  runExcelSnippetTest,
  runWordSnippetTest,
  runPowerPointSnippetTest,
  runCommonApiSnippetTest,
} from './helpers/snippet-test-runner';
import * as path from 'path';
import * as ts from 'typescript';
import { runInNewContext } from 'vm';
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

  test.each([false, true])('Excel: registers linked entity domains (existing: %s)', async (exists) => {
    const { mockContext } = await runExcelSnippetTest({
      snippetPath: path.join('samples', 'excel', '20-data-types', 'data-types-linked-entities.yaml'),
      buttonId: 'register-domains',
      mockOptions: { linkedEntityDomainExists: exists },
    });
    const domains = mockContext.workbook.linkedEntityDataDomains;
    expect(domains.add.mock.calls).toEqual([
      [{
        dataProvider: 'Contoso', id: 'products', name: 'Products',
        loadFunctionId: 'LINKEDENTITYLOADSERVICE.CONTOSOLOADSERVICE',
        periodicRefreshInterval: 300, supportedRefreshModes: ['Manual', 'OnLoad', 'Periodic'],
      }],
      [{
        dataProvider: 'Contoso', id: 'categories', name: 'Categories',
        loadFunctionId: 'LINKEDENTITYLOADSERVICE.CONTOSOLOADSERVICE',
        periodicRefreshInterval: 300, supportedRefreshModes: ['Manual', 'OnLoad', 'Periodic'],
      }],
      [{
        dataProvider: 'Contoso', id: 'suppliers', name: 'Suppliers',
        loadFunctionId: 'LINKEDENTITYLOADSERVICE.CONTOSOLOADSERVICE',
      }],
    ]);
    domains.items.forEach((domain) => expect(domain.delete).toHaveBeenCalledTimes(exists ? 1 : 0));
    expect(consoleErrorSpy).not.toHaveBeenCalled();
  });

  test('Excel: linked entity setup registers refresh notifications', async () => {
    const { mockContext } = await runExcelSnippetTest({
      snippetPath: path.join('samples', 'excel', '20-data-types', 'data-types-linked-entities.yaml'),
      buttonId: 'setup',
    });
    const domains = mockContext.workbook.linkedEntityDataDomains;
    expect(domains.onRefreshCompleted.add).toHaveBeenCalledTimes(1);
    const handler = domains.onRefreshCompleted.add.mock.calls[0][0];
    await handler({ id: 'products', serviceId: 268436224, refreshed: true });
    expect(consoleSpy).toHaveBeenCalledWith('Domain refreshed: products; succeeded: true');
    expect(consoleErrorSpy).not.toHaveBeenCalled();
    await handler({ id: 'products', serviceId: 268436224, refreshed: false, errors: ['Provider unavailable'] });
    expect(consoleErrorSpy).toHaveBeenCalledWith('Linked entity refresh errors:', ['Provider unavailable']);
  });

  test.each(['setup', 'insert-product', 'get-product', 'get-domains', 'refresh-product', 'refresh-domains'])(
    'Excel: linked entities %s action executes without runtime errors',
    async (buttonId) => {
      const { mockContext, mockRange, mockWorksheet } = await runExcelSnippetTest({
        snippetPath: path.join('samples', 'excel', '20-data-types', 'data-types-linked-entities.yaml'),
        buttonId,
        mockOptions: { rangeValuesAsJson: [[{
          type: 'LinkedEntity', text: 'Chai',
          id: { entityId: 'P1', domainId: 'products', serviceId: 268436224, culture: 'en-US' },
        }]] },
      });
      const domains = mockContext.workbook.linkedEntityDataDomains;
      if (buttonId === 'insert-product') {
        expect(mockWorksheet.getRange).toHaveBeenCalledWith('A2');
        expect(mockRange.valuesAsJson).toEqual([[{
          type: 'LinkedEntity', text: 'Chai', basicType: 'Error', basicValue: '#VALUE!',
          id: { entityId: 'P1', domainId: 'products', serviceId: 268436224, culture: 'en-US' },
        }]]);
        expect(mockRange.formulas).toEqual([['=A2.[Unit Price]', '=A2.Category', '=A2.Supplier']]);
      } else if (buttonId === 'get-product') {
        expect(mockRange.load).toHaveBeenCalledWith('valuesAsJson');
        expect(consoleSpy).toHaveBeenCalledWith('Chai linked entity:', mockRange.valuesAsJson[0][0]);
      } else if (buttonId === 'refresh-product') {
        expect(domains.getItem).toHaveBeenCalledWith('products');
        expect(domains.items[0].refresh).toHaveBeenCalledTimes(1);
      } else if (buttonId === 'refresh-domains') {
        expect(domains.refreshAll).toHaveBeenCalledTimes(1);
      } else if (buttonId === 'get-domains') {
        expect(domains.load).toHaveBeenCalled();
        expect(consoleSpy).toHaveBeenCalledWith('Domain settings:', expect.objectContaining({ id: 'products' }));
      } else {
        expect(mockContext.workbook.worksheets.add).toHaveBeenCalledWith('Sample');
        expect(mockWorksheet.activate).toHaveBeenCalledTimes(1);
      }
      expect(mockContext.sync).toHaveBeenCalled();
      expect(consoleErrorSpy).not.toHaveBeenCalled();
    }
  );

  test('Excel: linked entity inspection reports a missing product', async () => {
    await runExcelSnippetTest({
      snippetPath: path.join('samples', 'excel', '20-data-types', 'data-types-linked-entities.yaml'),
      buttonId: 'get-product',
    });
    expect(consoleErrorSpy).toHaveBeenCalledWith(expect.objectContaining({
      message: 'A2 is not a linked entity. Register the domains and insert the product first.',
    }));
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

  describe('Word: content control states (preview)', () => {
    const snippetPath = path.join('samples', 'word', '99-preview-apis', 'insert-and-change-content-controls.yaml');
    let wordMock: ReturnType<typeof createWordMock>;
    let buttonHandlers: Map<string, Function>;

    beforeEach(() => {
      wordMock = createWordMock();
      (global as any).Word = wordMock.mockObject;
      buttonHandlers = new Map<string, Function>();
      (global as any).document = createMockDocument(buttonHandlers);
      executeSnippetCode(loadSnippetByPath(snippetPath));
    });

    test('registers focused actions and recreates the tagged control during setup', async () => {
      expect([...buttonHandlers.keys()]).toEqual([
        'setup', 'set-warning-state', 'set-error-state', 'reset-state',
      ]);

      await clickButton(buttonHandlers, 'setup');
      await clickButton(buttonHandlers, 'setup');

      expect(wordMock.mockBody.clear).toHaveBeenCalledTimes(2);
      expect(wordMock.mockBody.insertParagraph.mock.calls).toEqual([
        ['Project code: pending review.', 'Start'],
        ['Project code: pending review.', 'Start'],
      ]);
      expect(wordMock.mockParagraph.getRange).toHaveBeenCalledWith('Content');
      expect(wordMock.mockRange.insertContentControl).toHaveBeenCalledTimes(2);
      const sampleControl = wordMock.mockRange.contentControls.items[0];
      expect(sampleControl.tag).toBe('ContentControlStateSample');
      expect(sampleControl.title).toBe('Project code');
      expect(sampleControl.select).toHaveBeenCalledTimes(2);
      expect(wordMock.mockContext.sync).toHaveBeenCalledTimes(2);
      expect(consoleErrorSpy).not.toHaveBeenCalled();
    });

    test.each([
      { buttonId: 'set-warning-state', state: 'Warning' },
      { buttonId: 'set-error-state', state: 'Error' },
      { buttonId: 'reset-state', state: null },
    ])('$buttonId targets the sample control', async ({ buttonId, state }) => {
      await clickButton(buttonHandlers, 'setup');
      consoleSpy.mockClear();

      await clickButton(buttonHandlers, buttonId);

      expect(wordMock.mockContext.document.contentControls.getByTag)
        .toHaveBeenCalledWith('ContentControlStateSample');
      const sampleControl = wordMock.mockRange.contentControls.items[0];
      if (state === null) {
        expect(sampleControl.resetState).toHaveBeenCalledTimes(1);
        expect(sampleControl.setState).not.toHaveBeenCalled();
        expect(consoleSpy).toHaveBeenCalledWith('Reset the sample content control state.');
      } else {
        expect(sampleControl.setState).toHaveBeenCalledWith(state);
        expect(sampleControl.resetState).not.toHaveBeenCalled();
        expect(consoleSpy).toHaveBeenCalledWith(`Set sample content control state to ${state}.`);
      }
      expect(wordMock.mockContext.sync).toHaveBeenCalledTimes(3);
      expect(consoleErrorSpy).not.toHaveBeenCalled();
    });

    test.each(['set-warning-state', 'set-error-state', 'reset-state'])(
      '%s warns when the sample is absent instead of changing an unrelated control',
      async (buttonId) => {
        const warningSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});

        await clickButton(buttonHandlers, buttonId);

        expect(wordMock.mockContext.document.contentControls.getByTag)
          .toHaveBeenCalledWith('ContentControlStateSample');
        const unrelatedControl = wordMock.mockRange.contentControls.items[0];
        expect(unrelatedControl.tag).toBe('RangeSampleSummary');
        expect(unrelatedControl.setState).not.toHaveBeenCalled();
        expect(unrelatedControl.resetState).not.toHaveBeenCalled();
        expect(warningSpy).toHaveBeenCalledWith('Run Set up to create the sample content control.');
        expect(consoleSpy).not.toHaveBeenCalled();
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      }
    );

    test.each(['set-warning-state', 'set-error-state', 'reset-state'])(
      '%s reports a failed Word request without logging success',
      async (buttonId) => {
        await clickButton(buttonHandlers, 'setup');
        consoleSpy.mockClear();
        const error = new Error('The Word client does not support content control states.');
        wordMock.mockContext.sync.mockResolvedValueOnce(undefined).mockRejectedValueOnce(error);

        await clickButton(buttonHandlers, buttonId);

        expect(consoleErrorSpy).toHaveBeenCalledWith(error);
        expect(consoleSpy).not.toHaveBeenCalled();
      }
    );
  });

  describe('Word: comment events (preview)', () => {
    const snippetPath = path.join('samples', 'word', '99-preview-apis', 'comment-events.yaml');
    let wordMock: ReturnType<typeof createWordMock>;
    let buttonHandlers: Map<string, Function>;
    let warningSpy: jest.SpyInstance;

    beforeEach(() => {
      wordMock = createWordMock({ bookmarkNames: ['CommentEventsSample'] });
      (global as any).Word = wordMock.mockObject;
      buttonHandlers = new Map<string, Function>();
      (global as any).document = createMockDocument(buttonHandlers);
      warningSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
      executeSnippetCode(loadSnippetByPath(snippetPath));
    });

    test('exposes only setup, event registration, and focused event triggers', () => {
      expect([...buttonHandlers.keys()]).toEqual([
        'setup', 'register-event-handlers', 'create-comment', 'edit-comment',
        'delete-comment', 'deregister-event-handlers',
      ]);
    });

    const commentEvents = [
      'onCommentAdded', 'onCommentChanged', 'onCommentDeleted', 'onCommentDeselected', 'onCommentSelected',
    ] as const;

    test.each(commentEvents)('%s waits for body tracking to synchronize before subscribing', async (eventName) => {
      let completeTracking: () => void;
      let notifyTrackingStarted: () => void;
      const trackingSync = new Promise<void>(resolve => { completeTracking = resolve; });
      const trackingStarted = new Promise<void>(resolve => { notifyTrackingStarted = resolve; });
      wordMock.mockBody.track.mockImplementation(() => {
        if (wordMock.mockBody.track.mock.calls.length === commentEvents.indexOf(eventName) + 1) {
          wordMock.mockContext.sync.mockImplementationOnce(() => {
            notifyTrackingStarted();
            return trackingSync;
          });
        }
      });

      const registration = clickButton(buttonHandlers, 'register-event-handlers');
      await trackingStarted;
      try {
        expect(wordMock.mockCommentEvents[eventName].add).not.toHaveBeenCalled();
      } finally {
        completeTracking();
        await registration;
      }

      expect(wordMock.mockCommentEvents[eventName].add).toHaveBeenCalledTimes(1);
      expect(consoleErrorSpy).not.toHaveBeenCalled();
    });

    test('setup recreates the bookmarked paragraph and removes existing subscriptions', async () => {
      await clickButton(buttonHandlers, 'setup');
      await clickButton(buttonHandlers, 'register-event-handlers');
      const handlers = Object.values(wordMock.mockCommentEvents).map(event => event.add.mock.results[0].value);

      await clickButton(buttonHandlers, 'setup');

      expect(wordMock.mockBody.clear).toHaveBeenCalledTimes(2);
      expect(wordMock.mockBody.insertParagraph.mock.calls).toEqual([
        ['The project launches on October 30.', 'Start'],
        ['The project launches on October 30.', 'Start'],
      ]);
      expect(wordMock.mockParagraph.getRange).toHaveBeenCalledWith('Content');
      expect(wordMock.mockRange.insertBookmark.mock.calls).toEqual([
        ['CommentEventsSample'], ['CommentEventsSample'],
      ]);
      handlers.forEach(handler => expect(handler.remove).toHaveBeenCalledTimes(1));
      expect(wordMock.mockBody.untrack).toHaveBeenCalledTimes(5);
      expect(consoleErrorSpy).not.toHaveBeenCalled();
    });

    test('registering twice does not duplicate subscriptions, and handlers can be registered again after removal', async () => {
      await clickButton(buttonHandlers, 'register-event-handlers');
      await clickButton(buttonHandlers, 'register-event-handlers');

      const events = Object.values(wordMock.mockCommentEvents);
      events.forEach(event => expect(event.add).toHaveBeenCalledTimes(1));
      expect(wordMock.mockBody.track).toHaveBeenCalledTimes(5);
      expect(warningSpy).toHaveBeenCalledWith('Deregister the existing handlers before registering again.');
      const handlers = events.map(event => event.add.mock.results[0].value);

      await clickButton(buttonHandlers, 'deregister-event-handlers');
      await clickButton(buttonHandlers, 'deregister-event-handlers');

      expect(consoleSpy).toHaveBeenCalledWith('Number of event handlers to remove: 5');
      expect(consoleSpy).toHaveBeenCalledWith('Number of event handlers to remove: 0');
      handlers.forEach(handler => expect(handler.remove).toHaveBeenCalledTimes(1));
      expect(wordMock.mockBody.untrack).toHaveBeenCalledTimes(5);
      expect(wordMock.mockContext.sync).toHaveBeenCalledTimes(15);

      await clickButton(buttonHandlers, 'register-event-handlers');
      events.forEach(event => expect(event.add).toHaveBeenCalledTimes(2));
      expect(consoleErrorSpy).not.toHaveBeenCalled();
    });

    test.each([
      'CommentAdded', 'CommentChanged', 'CommentDeleted', 'CommentDeselected', 'CommentSelected',
    ] as const)('%s callbacks report structured event details without a Word request', async (type) => {
      await clickButton(buttonHandlers, 'register-event-handlers');
      consoleSpy.mockClear();
      wordMock.mockContext.sync.mockClear();
      const event = wordMock.mockCommentEvents[`on${type}`];
      const handler = event.add.mock.calls[0][0];

      await handler({
        type,
        source: 'Local',
        changeType: type === 'CommentChanged' ? 'edited' : 'none',
        commentDetails: [
          { id: 'comment-1', replyIds: ['reply-1'] },
          { id: 'comment-2', replyIds: [] },
        ],
      });

      expect(consoleSpy).toHaveBeenCalledWith(`${type} event. Source: Local.`);
      expect(consoleSpy).toHaveBeenCalledWith('Affected comment:', { id: 'comment-1', replyIds: ['reply-1'] });
      expect(consoleSpy).toHaveBeenCalledWith('Affected comment:', { id: 'comment-2', replyIds: [] });
      if (type === 'CommentChanged') {
        expect(consoleSpy).toHaveBeenCalledWith('Comment change: edited.');
        expect(consoleSpy).toHaveBeenCalledWith('The comment text was edited.');
      }
      expect(wordMock.mockContext.sync).not.toHaveBeenCalled();
      expect(consoleErrorSpy).not.toHaveBeenCalled();
    });

    test.each(['resolved', 'reopened', 'replyAdded', 'replyDeleted', 'replyEdited'])(
      'changed callbacks preserve the %s change type and remote source',
      async (changeType) => {
        await clickButton(buttonHandlers, 'register-event-handlers');
        consoleSpy.mockClear();
        const handler = wordMock.mockCommentEvents.onCommentChanged.add.mock.calls[0][0];

        await handler({
          type: 'CommentChanged',
          source: 'Remote',
          changeType,
          commentDetails: [{ id: 'comment-1', replyIds: ['reply-1'] }],
        });

        expect(consoleSpy).toHaveBeenCalledWith(`Comment change: ${changeType}.`);
        expect(consoleSpy).toHaveBeenCalledWith('CommentChanged event. Source: Remote.');
        expect(consoleSpy).toHaveBeenCalledWith('A coauthor triggered this event.');
        expect(consoleSpy).not.toHaveBeenCalledWith('The comment text was edited.');
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      }
    );

    test('failed registration reports the error and retains the handle for cleanup', async () => {
      const error = new Error('Comment event registration failed.');
      wordMock.mockContext.sync.mockResolvedValueOnce(undefined).mockRejectedValueOnce(error);

      await clickButton(buttonHandlers, 'register-event-handlers');

      expect(consoleErrorSpy).toHaveBeenCalledWith(error);
      expect(consoleSpy).not.toHaveBeenCalled();
      expect(wordMock.mockCommentEvents.onCommentChanged.add).not.toHaveBeenCalled();
      const handler = wordMock.mockCommentEvents.onCommentAdded.add.mock.results[0].value;

      await clickButton(buttonHandlers, 'deregister-event-handlers');

      expect(consoleSpy).toHaveBeenCalledWith('Number of event handlers to remove: 1');
      expect(handler.remove).toHaveBeenCalledTimes(1);
      expect(wordMock.mockBody.untrack).toHaveBeenCalledTimes(1);
    });

    test('failed body tracking prevents subscriptions and success messages', async () => {
      const error = new Error('Body tracking failed.');
      wordMock.mockContext.sync.mockRejectedValueOnce(error);

      await clickButton(buttonHandlers, 'register-event-handlers');

      Object.values(wordMock.mockCommentEvents).forEach(event => expect(event.add).not.toHaveBeenCalled());
      expect(consoleErrorSpy).toHaveBeenCalledWith(error);
      expect(consoleSpy).not.toHaveBeenCalled();
    });

    test('a failed deregistration stops setup before replacing the document', async () => {
      await clickButton(buttonHandlers, 'register-event-handlers');
      consoleSpy.mockClear();
      const error = new Error('Comment event removal failed.');
      wordMock.mockContext.sync.mockRejectedValueOnce(error);

      await clickButton(buttonHandlers, 'setup');

      expect(consoleErrorSpy).toHaveBeenCalledWith(error);
      expect(wordMock.mockBody.clear).not.toHaveBeenCalled();
      expect(consoleSpy).toHaveBeenCalledWith('Number of event handlers to remove: 5');
      expect(consoleSpy).not.toHaveBeenCalledWith('Removed comment event handlers.');
      expect(consoleSpy).not.toHaveBeenCalledWith('Created the sample paragraph. Register handlers before triggering events.');
      await clickButton(buttonHandlers, 'register-event-handlers');
      expect(warningSpy).toHaveBeenCalledWith('Deregister the existing handlers before registering again.');
      Object.values(wordMock.mockCommentEvents).forEach(event => expect(event.add).toHaveBeenCalledTimes(1));
    });

    test('creation targets the bookmarked paragraph and guards against duplicate comments', async () => {
      await clickButton(buttonHandlers, 'create-comment');

      expect(warningSpy).toHaveBeenCalledWith('The sample paragraph already has a comment. Edit or delete it first.');
      expect(wordMock.mockRange.insertComment).not.toHaveBeenCalled();

      wordMock.mockComment.isNullObject = true;
      await clickButton(buttonHandlers, 'create-comment');

      expect(wordMock.mockContext.document.getBookmarkRangeOrNullObject).toHaveBeenCalledWith('CommentEventsSample');
      expect(wordMock.mockRange.insertComment).toHaveBeenCalledWith('Please verify the launch date.');
      expect(wordMock.mockContext.document.getSelection).not.toHaveBeenCalled();
      expect(consoleErrorSpy).not.toHaveBeenCalled();
    });

    test('editing switches comment text on each click and deleting targets the same scoped comment', async () => {
      await clickButton(buttonHandlers, 'edit-comment');
      expect(wordMock.mockComment.load).toHaveBeenCalledWith('content,resolved');
      expect(wordMock.mockComment.content).toBe('The launch date is confirmed.');

      await clickButton(buttonHandlers, 'edit-comment');
      expect(wordMock.mockComment.content).toBe('Please verify the launch date.');

      await clickButton(buttonHandlers, 'delete-comment');
      expect(wordMock.mockComment.delete).toHaveBeenCalledTimes(1);
      expect(wordMock.mockRange.getComments).toHaveBeenCalledTimes(3);
      expect(wordMock.mockContext.document.getSelection).not.toHaveBeenCalled();
      expect(consoleErrorSpy).not.toHaveBeenCalled();
    });

    test.each(['create-comment', 'edit-comment', 'delete-comment'])(
      '%s warns when the sample bookmark is missing',
      async (buttonId) => {
        wordMock.mockContext.document.getBookmarkRangeOrNullObject.mockReturnValueOnce({
          isNullObject: true, load: jest.fn(),
        });

        await clickButton(buttonHandlers, buttonId);

        expect(warningSpy).toHaveBeenCalledWith('Run Set up to create the sample paragraph.');
        expect(wordMock.mockRange.getComments).not.toHaveBeenCalled();
        expect(wordMock.mockRange.insertComment).not.toHaveBeenCalled();
        expect(wordMock.mockComment.content).toBe('Please verify the launch date.');
        expect(wordMock.mockComment.delete).not.toHaveBeenCalled();
        expect(consoleSpy).not.toHaveBeenCalled();
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      }
    );

    test.each([
      ['edit-comment', 'editing'],
      ['delete-comment', 'deleting'],
    ])('%s warns when the scoped comment is absent', async (buttonId, action) => {
      wordMock.mockComment.isNullObject = true;

      await clickButton(buttonHandlers, buttonId);

      expect(warningSpy).toHaveBeenCalledWith(`Create the sample comment before ${action} it.`);
      expect(wordMock.mockComment.content).toBe('Please verify the launch date.');
      expect(wordMock.mockComment.delete).not.toHaveBeenCalled();
      expect(consoleSpy).not.toHaveBeenCalled();
      expect(consoleErrorSpy).not.toHaveBeenCalled();
    });

    test('editing a resolved comment warns instead of attempting a change', async () => {
      wordMock.mockComment.resolved = true;

      await clickButton(buttonHandlers, 'edit-comment');

      expect(warningSpy).toHaveBeenCalledWith('Reopen the sample comment in Word before editing it.');
      expect(wordMock.mockComment.content).toBe('Please verify the launch date.');
      expect(consoleSpy).not.toHaveBeenCalled();
      expect(consoleErrorSpy).not.toHaveBeenCalled();
    });

    test.each(['create-comment', 'edit-comment', 'delete-comment'])(
      '%s reports a failed mutation without logging success',
      async (buttonId) => {
        wordMock.mockComment.isNullObject = buttonId === 'create-comment';
        const error = new Error('Comment operation failed.');
        wordMock.mockContext.sync.mockResolvedValueOnce(undefined).mockResolvedValueOnce(undefined).mockRejectedValueOnce(error);

        await clickButton(buttonHandlers, buttonId);

        expect(consoleErrorSpy).toHaveBeenCalledWith(error);
        expect(consoleSpy).not.toHaveBeenCalled();
      }
    );
  });

  test.each([true, false])('Word: close document window checks desktop support (%s)', async (supported) => {
    const isSetSupported = jest.fn(() => supported);
    (global as any).Office = { context: { requirements: { isSetSupported } } };

    const { mockContext } = await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '50-document', 'close-document-window.yaml'),
      buttonId: 'close-with-save-prompt',
    });

    expect(isSetSupported).toHaveBeenCalledWith('WordApiDesktop', '1.4');
    if (supported) {
      expect(mockContext.document.activeWindow.close).toHaveBeenCalledWith({
        saveChanges: 'PromptToSaveChanges',
      });
      expect(mockContext.sync).toHaveBeenCalledTimes(1);
      expect(consoleSpy).toHaveBeenCalledWith('About to close the document window...');
      expect(consoleErrorSpy).not.toHaveBeenCalled();
    } else {
      expect(mockContext.document.activeWindow.close).not.toHaveBeenCalled();
      expect(mockContext.sync).not.toHaveBeenCalled();
      expect(consoleSpy).not.toHaveBeenCalled();
      expect(consoleErrorSpy).toHaveBeenCalledWith('This sample requires WordApiDesktop 1.4 or later.');
    }
  });

  test('Word: close document window reports a failed close request', async () => {
    const { mockObject, mockContext } = createWordMock();
    const error = new Error('The document window could not be closed.');
    mockContext.sync.mockRejectedValueOnce(error);
    (global as any).Office = { context: { requirements: { isSetSupported: jest.fn(() => true) } } };
    (global as any).Word = mockObject;
    const buttonHandlers = new Map<string, Function>();
    (global as any).document = createMockDocument(buttonHandlers);
    executeSnippetCode(loadSnippetByPath(path.join('samples', 'word', '50-document', 'close-document-window.yaml')));

    await clickButton(buttonHandlers, 'close-with-save-prompt');

    expect(mockContext.document.activeWindow.close).toHaveBeenCalledWith({
      saveChanges: 'PromptToSaveChanges',
    });
    expect(consoleErrorSpy).toHaveBeenCalledWith(error);
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

  test('Word: table structure setup replaces the document with a sample table', async () => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '40-tables', 'manage-table-structure.yaml'),
      buttonId: 'setup',
      assertions: ({ mockContext, mockBody }) => {
        expect(mockBody.clear).toHaveBeenCalled();
        expect(mockBody.insertTable).toHaveBeenCalledWith(4, 3, 'Start', [
          ['Region', 'Sales', 'Status'],
          ['North', '$125,000', 'On track'],
          ['South', '$98,000', 'Needs review'],
          ['West', '$143,000', 'On track'],
        ]);
        expect(mockContext.sync).toHaveBeenCalled();
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    { buttonId: 'add-columns', method: 'addColumns', args: ['End', 1] },
    { buttonId: 'add-rows', method: 'addRows', args: ['End', 1] },
    { buttonId: 'delete-columns', method: 'deleteColumns', args: [2, 1] },
    { buttonId: 'delete-rows', method: 'deleteRows', args: [1, 1] },
    { buttonId: 'insert-paragraph', method: 'insertParagraph', args: ['Quarterly sales', 'Before'] },
    {
      buttonId: 'insert-table',
      method: 'insertTable',
      args: [2, 2, 'After', [['Item', 'Note'], ['Sales', 'Review quarterly totals.']]],
    },
  ])('Word: table structure $buttonId uses the Table method', async ({ buttonId, method, args }) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '40-tables', 'manage-table-structure.yaml'),
      buttonId,
      assertions: ({ mockContext }) => {
        const table = mockContext.document.body.tables.getFirst();
        expect(table[method]).toHaveBeenCalledWith(...args);
        expect(mockContext.sync).toHaveBeenCalled();
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    {
      buttonId: 'delete-columns',
      warning: 'Keep at least one column. Add a column or choose Set up to reset the sample.',
    },
    {
      buttonId: 'delete-rows',
      warning: 'Keep the header row. Add a row or choose Set up to reset the sample.',
    },
  ])('Word: table structure $buttonId preserves the final row or column', async ({ buttonId, warning }) => {
    const { mockObject, mockTable } = createWordMock();
    Object.assign(mockTable, { rowCount: 1 });
    mockTable.rows.getFirst().cellCount = 1;
    (global as any).Word = mockObject;
    const buttonHandlers = new Map<string, Function>();
    (global as any).document = createMockDocument(buttonHandlers);
    const warningSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
    executeSnippetCode(loadSnippetByPath(path.join('samples', 'word', '40-tables', 'manage-table-structure.yaml')));

    await clickButton(buttonHandlers, buttonId);

    expect(mockTable.deleteColumns).not.toHaveBeenCalled();
    expect(mockTable.deleteRows).not.toHaveBeenCalled();
    expect(warningSpy).toHaveBeenCalledWith(warning);
    expect(consoleErrorSpy).not.toHaveBeenCalled();
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

describe('Runtime Execution Tests - Linked entity load service', () => {
  test('uses the linked entity load service registration signature', () => {
    const snippet = loadSnippetByPath(path.join('samples', 'excel', '20-data-types', 'data-types-linked-entity-load-service.yaml'));
    const source = ts.createSourceFile('snippet.ts', snippet.script.content, ts.ScriptTarget.Latest, true);
    const loadService = source.statements.find((statement): statement is ts.FunctionDeclaration =>
      ts.isFunctionDeclaration(statement) && statement.name?.text === 'contosoLoadService'
    );
    if (!loadService) {
      throw new Error('The linked entity custom function was not found.');
    }
    const tags = ts.getJSDocTags(loadService).map((tag) => tag.tagName.text.toLowerCase());
    expect(tags).toContain('linkedentityloadservice');
    expect(tags).not.toContain('linkedentitydataprovider');
    expect(loadService.parameters).toHaveLength(1);
    const requestParameter = loadService.parameters[0];
    expect(requestParameter.type?.kind).toBe(ts.SyntaxKind.UnknownKeyword);
    expect(ts.getJSDocType(requestParameter)?.kind).toBe(ts.SyntaxKind.UnknownKeyword);
    expect(requestParameter.questionToken).toBeUndefined();
    expect(requestParameter.dotDotDotToken).toBeUndefined();
    expect(requestParameter.initializer).toBeUndefined();
    const returnType = ts.getJSDocReturnType(loadService);
    expect(returnType?.kind).toBe(ts.SyntaxKind.AnyKeyword);
    // Script Lab compares explicit TypeScript and JSDoc return syntax kinds, not their assignability.
    expect(loadService.type?.kind ?? returnType?.kind).toBe(returnType?.kind);
  });

  function createLoadService() {
    const snippet = loadSnippetByPath(path.join('samples', 'excel', '20-data-types', 'data-types-linked-entity-load-service.yaml'));
    const code = ts.transpileModule(`${snippet.script.content}\ncontosoLoadService;`, {
      compilerOptions: { target: ts.ScriptTarget.ES2024, module: ts.ModuleKind.None },
    }).outputText;
    class CustomFunctionError extends Error {
      constructor(public code: string, message: string) {
        super(message);
      }
    }
    const errorSpy = jest.fn();
    const loadService = runInNewContext(code, {
      console: { log: jest.fn(), error: errorSpy },
      CustomFunctions: { Error: CustomFunctionError, ErrorCode: { notAvailable: '#N/A' } },
    }) as (request: unknown) => {
      entities: Array<{
        id: { entityId: string; domainId: string; serviceId: number; culture: string };
        properties: Record<string, unknown>;
      }>;
    };
    return { loadService, errorSpy };
  }

  test('returns a product card with deferred category and supplier links', () => {
    const { loadService, errorSpy } = createLoadService();
    const response = loadService(JSON.stringify({
      domainId: 'products', entities: [{ entityId: 'P1', culture: 'en-US' }],
    }));
    expect(response.entities).toHaveLength(1);
    expect(response.entities[0]).toMatchObject({
      type: 'LinkedEntity', text: 'Chai', basicType: 'Error', basicValue: '#VALUE!',
      id: { entityId: 'P1', domainId: 'products', serviceId: 268436224, culture: 'en-US' },
      properties: { 'Unit Price': { type: 'FormattedNumber', basicValue: 18 } },
      layouts: { card: { title: { property: 'Product Name' } } },
    });
    expect(response.entities[0].properties.Category).toEqual({
      type: 'LinkedEntity', text: 'Beverages', basicType: 'Error', basicValue: '#VALUE!',
      id: { entityId: 'C1', domainId: 'categories', serviceId: 268436224, culture: 'en-US' },
    });
    expect(response.entities[0].properties.Supplier).toEqual({
      type: 'LinkedEntity', text: 'Exotic Liquids', basicType: 'Error', basicValue: '#VALUE!',
      id: { entityId: 'S1', domainId: 'suppliers', serviceId: 268436224, culture: 'en-US' },
    });
    expect(errorSpy).not.toHaveBeenCalled();
  });

  test.each([
    { domainId: 'products', entityId: 'P1', property: 'Product Name', value: 'Chai' },
    { domainId: 'categories', entityId: 'C1', property: 'Category Name', value: 'Beverages' },
    { domainId: 'suppliers', entityId: 'S1', property: 'Contact Name', value: 'Ema Vargova' },
  ])('resolves $domainId in request order and preserves cultures', ({ domainId, entityId, property, value }) => {
    const { loadService } = createLoadService();
    const response = loadService(JSON.stringify({
      domainId, entities: [{ entityId, culture: 'fr-FR' }, { entityId, culture: 'en-US' }],
    }));
    expect(response.entities.map((entity) => entity.id)).toEqual([
      { entityId, domainId, serviceId: 268436224, culture: 'fr-FR' },
      { entityId, domainId, serviceId: 268436224, culture: 'en-US' },
    ]);
    response.entities.forEach((entity) => expect(entity.properties[property]).toEqual({ type: 'String', basicValue: value }));
  });

  test.each([
    undefined,
    null,
    42,
    { domainId: 'products', entities: [{ entityId: 'P1', culture: 'en-US' }] },
    'not JSON',
    'null',
    '{}',
    JSON.stringify({ domainId: 'products', entities: [null] }),
    JSON.stringify({ domainId: 'products', entities: [{ entityId: 'P1' }] }),
    JSON.stringify({ domainId: 'unknown', entities: [{ entityId: 'P1', culture: 'en-US' }] }),
    JSON.stringify({ domainId: 'products', entities: [{ entityId: 'missing', culture: 'en-US' }] }),
    JSON.stringify({ domainId: 'categories', entities: [{ entityId: 'missing', culture: 'en-US' }] }),
    JSON.stringify({ domainId: 'suppliers', entities: [{ entityId: 'missing', culture: 'en-US' }] }),
  ])('reports an unavailable error for invalid request %s', (request) => {
    const { loadService, errorSpy } = createLoadService();
    expect(() => loadService(request)).toThrow('Unable to load the requested Contoso entity.');
    expect(() => loadService(request)).toThrow(expect.objectContaining({ code: '#N/A' }));
    expect(errorSpy).toHaveBeenCalledWith('Unable to load linked entities:', expect.anything());
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
