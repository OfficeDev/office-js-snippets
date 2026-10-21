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

  test.each([
    ['LabelingEnabled', 'Sensitivity labeling is enabled.'],
    ['LabelingDisabled', 'Sensitivity labeling is disabled.'],
    ['LabelingPolicyNotFound', 'No sensitivity labeling policy was found.'],
    ['NoLicense', "The current user doesn't have a valid sensitivity labeling license."],
    ['FutureCapability', 'Unknown labeling capability: FutureCapability'],
  ])('Word: catalog status %s displays the expected message', async (labelingCapability, expectedMessage) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '99-preview-apis', 'manage-sensitivity-labels.yaml'),
      buttonId: 'get-catalog-status',
      mockOptions: { labelingCapability },
      assertions: ({ mockContext }) => {
        expect(mockContext.sensitivityLabelsCatalog.load).toHaveBeenCalledWith('getLabelingCapability');
        expect(mockContext.sync).toHaveBeenCalledTimes(1);
        expect(consoleSpy).toHaveBeenCalledWith(expectedMessage);
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    { abacEnabled: true, isNullObject: false },
    { abacEnabled: false, isNullObject: false },
    { abacEnabled: true, isNullObject: true },
  ])('Word: getting a label handles ABAC $abacEnabled and null label $isNullObject', async ({
    abacEnabled, isNullObject,
  }) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '99-preview-apis', 'manage-sensitivity-labels.yaml'),
      buttonId: 'get-document-label',
      mockOptions: {
        sensitivityLabelAbacEnabled: abacEnabled,
        sensitivityLabelIsNullObject: isNullObject,
      },
      assertions: ({ mockContext }) => {
        const getCurrentLabel = mockContext.document.sensitivityLabel.getCurrentOrNullObject;
        expect(getCurrentLabel).toHaveBeenCalledTimes(1);
        const currentLabel = getCurrentLabel.mock.results[0].value;
        expect(currentLabel.load).toHaveBeenCalledWith(
          'color,id,isAbacAttributesRequired,isAbacEnabled,isEnabled,name,order,protectionType,siteId,tooltip'
        );
        if (isNullObject) {
          expect(consoleSpy).toHaveBeenCalledWith("The document doesn't have a sensitivity label.");
        } else if (abacEnabled) {
          expect(currentLabel.load).toHaveBeenCalledWith('abacAttributeValues');
          expect(consoleSpy).toHaveBeenCalledWith(expect.stringContaining(
            'ABAC attribute values: Engineering (ID: engineering, order: 1, active: Yes), ' +
            'Finance (ID: finance, order: 2, active: No)'
          ));
        } else {
          expect(currentLabel.load).not.toHaveBeenCalledWith('abacAttributeValues');
          expect(consoleSpy).toHaveBeenCalledWith(expect.stringContaining('ABAC attribute values: None'));
        }
        expect(mockContext.sync).toHaveBeenCalledTimes(abacEnabled && !isNullObject ? 2 : 1);
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      },
    });
  });

  test.each([
    { updateResult: 'Success', successful: false },
    {
      updateResult: {
        '@odata.type': 'Microsoft.WordServices.SensitivityLabelUpdateResult',
        status: 'Success',
      },
      successful: true,
    },
    { updateResult: 'DocumentReadOnly', successful: false },
    { updateResult: { status: 'DocumentReadOnly' }, successful: false },
    { updateResult: { unexpected: 'Success' }, successful: false },
    { updateResult: null, successful: false },
  ])('Word: setting a label handles update result $updateResult', async ({ updateResult, successful }) => {
    await runWordSnippetTest({
      snippetPath: path.join('samples', 'word', '99-preview-apis', 'manage-sensitivity-labels.yaml'),
      buttonId: 'set-document-label',
      elementValues: { 'label-id': 'sample-label' },
      mockOptions: { sensitivityLabelUpdateResult: updateResult },
      assertions: ({ mockContext }) => {
        expect(mockContext.document.sensitivityLabel.tryToUpdate).toHaveBeenCalledWith('sample-label');
        expect(mockContext.sync).toHaveBeenCalledTimes(1);
        if (successful) {
          expect(consoleSpy).toHaveBeenCalledWith('Applied the sensitivity label to the document.');
          expect(consoleErrorSpy).not.toHaveBeenCalled();
        } else {
          expect(consoleErrorSpy).toHaveBeenCalledWith(
            `Sensitivity label update returned a non-success result: ${JSON.stringify(updateResult)}`
          );
          expect(consoleSpy).not.toHaveBeenCalled();
        }
      },
    });
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

  test.each([
    { updateResult: { status: 'Success' }, successful: true },
    { updateResult: { status: 'DocumentReadOnly' }, successful: false },
    { updateResult: 'Success', successful: false },
    { updateResult: { unexpected: 'Success' }, successful: false },
    { updateResult: null, successful: false },
  ])('PowerPoint: setting a label handles update result $updateResult without the runtime enum', async ({ updateResult, successful }) => {
    await runPowerPointSnippetTest({
      snippetPath: path.join('samples', 'powerpoint', 'preview-apis', 'manage-sensitivity-labels.yaml'),
      buttonId: 'set-presentation-label',
      elementValues: { 'label-id': 'sample-label' },
      mockOptions: { sensitivityLabelUpdateResult: updateResult },
      assertions: ({ mockContext }) => {
        expect(mockContext.presentation.sensitivityLabel.tryToUpdate).toHaveBeenCalledWith('sample-label');
        expect(mockContext.sync).toHaveBeenCalledTimes(1);
        if (successful) {
          expect(consoleSpy).toHaveBeenCalledWith('Applied the sensitivity label to the presentation.');
          expect(consoleErrorSpy).not.toHaveBeenCalled();
        } else {
          expect(consoleErrorSpy).toHaveBeenCalledWith(
            `Sensitivity label update returned a non-success result: ${JSON.stringify(updateResult)}`
          );
          expect(consoleSpy).not.toHaveBeenCalled();
        }
      },
    });
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
