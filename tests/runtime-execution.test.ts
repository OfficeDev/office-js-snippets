import {
  runExcelSnippetTest,
  runWordSnippetTest,
  runPowerPointSnippetTest,
  runCommonApiSnippetTest,
} from './helpers/snippet-test-runner';
import * as path from 'path';
import { clickButton } from './helpers/snippet-runtime';

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

  const sensitivitySnippetPath = path.join('samples', 'excel', '98-preview-apis', 'manage-sensitivity-labels.yaml');

  test.each([
    ['LabelingEnabled', 'Sensitivity labeling is enabled.'],
    ['LabelingDisabled', 'Sensitivity labeling is disabled.'],
    ['LabelingPolicyNotFound', 'No sensitivity labeling policy was found.'],
    ['NoLicense', "The current user doesn't have a valid sensitivity labeling license."],
    ['AttributeBasedAccessControlEnabled', 'Labeling capability: AttributeBasedAccessControlEnabled'],
  ])('Excel: catalog status %s displays the expected message', async (labelingCapability, expectedMessage) => {
    const { mockContext } = await runExcelSnippetTest({
      snippetPath: sensitivitySnippetPath,
      buttonId: 'get-catalog-status',
      mockOptions: { labelingCapability },
    });
    expect(mockContext.sensitivityLabelsCatalog?.load).toHaveBeenCalledWith('getLabelingCapability');
    expect(mockContext.sync).toHaveBeenCalledTimes(1);
    expect(consoleSpy).toHaveBeenCalledWith(expectedMessage);
    expect(consoleErrorSpy).not.toHaveBeenCalled();
  });

  test.each(['get-catalog-status', 'get-available-labels', 'get-catalog-attributes'])(
    'Excel: %s reports an unavailable catalog',
    async (buttonId) => {
      const warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
      const { mockContext } = await runExcelSnippetTest({
        snippetPath: sensitivitySnippetPath,
        buttonId,
        mockOptions: { sensitivityLabelCatalogAvailable: false },
      });
      expect(warnSpy).toHaveBeenCalledWith(
        "The sensitivity label catalog isn't available in this version of Excel."
      );
      expect(mockContext.sync).not.toHaveBeenCalled();
      expect(consoleErrorSpy).not.toHaveBeenCalled();
    }
  );

  test.each(['populated', 'empty', 'null'])('Excel: catalog attributes handle %s values', async (valueState) => {
    const { mockContext } = await runExcelSnippetTest({
      snippetPath: sensitivitySnippetPath,
      buttonId: 'get-catalog-attributes',
      mockOptions: {
        sensitivityLabelAttributeValuesEmpty: valueState === 'empty',
        sensitivityLabelAttributeValuesNull: valueState === 'null',
      },
    });
    const getAttributes = mockContext.sensitivityLabelsCatalog!.getAttributes;
    expect(getAttributes).toHaveBeenCalledTimes(1);
    const attribute = getAttributes.mock.results[0].value.value[0];
    expect(attribute.load).toHaveBeenCalledWith(
      'dataType,displayName,displayOrder,isActive,isMultiValued,values'
    );
    attribute.values?.forEach(attributeValue => {
      expect(attributeValue.load).toHaveBeenCalledWith('displayName,displayOrder,id,isActive');
    });
    expect(mockContext.sync).toHaveBeenCalledTimes(3);
    expect(consoleSpy).toHaveBeenCalledWith(
      '\nDepartment\n  Data type: StringValue\n  Display order: 1\n  Active: Yes\n  Multiple values: No\n  Values: ' +
      (valueState === 'populated'
        ? 'Engineering (ID: engineering, order: 1, active: Yes), Finance (ID: finance, order: 2, active: No)'
        : 'None')
    );
    expect(consoleErrorSpy).not.toHaveBeenCalled();
  });

  test.each(['empty', 'null'])('Excel: a %s attribute catalog reports no attributes', async (catalogState) => {
    const { mockContext } = await runExcelSnippetTest({
      snippetPath: sensitivitySnippetPath,
      buttonId: 'get-catalog-attributes',
      mockOptions: {
        sensitivityLabelAttributesEmpty: catalogState === 'empty',
        sensitivityLabelAttributesNull: catalogState === 'null',
      },
    });
    expect(mockContext.sensitivityLabelsCatalog?.getAttributes).toHaveBeenCalledTimes(1);
    expect(mockContext.sync).toHaveBeenCalledTimes(1);
    expect(consoleSpy).toHaveBeenCalledWith('No ABAC attributes are available in the sensitivity label catalog.');
    expect(consoleErrorSpy).not.toHaveBeenCalled();
  });

  test('Excel: available labels populate enabled options and grouped sublabels without duplicates', async () => {
    const { mockContext, mockDocument, buttonHandlers } = await runExcelSnippetTest({
      snippetPath: sensitivitySnippetPath,
      buttonId: 'get-available-labels',
    });
    await clickButton(buttonHandlers, 'get-available-labels');
    expect(mockContext.sensitivityLabelsCatalog?.getLabels).toHaveBeenCalledTimes(2);
    expect(mockContext.sensitivityLabelsCatalog?.getLabels.mock.results[0].value.load).toHaveBeenCalledWith(
      'items/color,items/id,items/isEnabled,items/name,items/order,items/protectionType,' +
      'items/siteId,items/tooltip,items/children/items/id,items/children/items/name,items/children/items/isEnabled'
    );
    const options = mockDocument.getElementById('label-id').options;
    expect(options).toHaveLength(4);
    expect(options[0].value).toBe('');
    expect(options[1]).toMatchObject({ value: 'general-label', disabled: false });
    expect(options[2]).toMatchObject({ value: 'disabled-label', disabled: true });
    expect(options[3].label).toBe('Confidential');
    expect(options[3].children).toEqual([
      expect.objectContaining({ value: 'engineering-label', disabled: false }),
      expect.objectContaining({ value: 'retired-label', disabled: true }),
    ]);
    expect(options.some(option => option.value === 'parent-label')).toBe(false);
    expect(consoleSpy).toHaveBeenCalledWith(expect.stringContaining('Engineering (engineering-label)'));
    expect(consoleErrorSpy).not.toHaveBeenCalled();
  });

  test('Excel: an empty label catalog leaves only the placeholder', async () => {
    const { mockDocument } = await runExcelSnippetTest({
      snippetPath: sensitivitySnippetPath,
      buttonId: 'get-available-labels',
      mockOptions: { sensitivityLabelsEmpty: true },
    });
    expect(mockDocument.getElementById('label-id').options).toHaveLength(1);
    expect(consoleSpy).toHaveBeenCalledWith('No sensitivity labels are available for the current user.');
    expect(consoleErrorSpy).not.toHaveBeenCalled();
  });

  test.each([false, true])('Excel: getting a workbook label handles null label: %s', async (isNullObject) => {
    const { mockContext } = await runExcelSnippetTest({
      snippetPath: sensitivitySnippetPath,
      buttonId: 'get-workbook-label',
      mockOptions: { sensitivityLabelIsNullObject: isNullObject },
    });
    const getCurrentLabel = mockContext.workbook.sensitivityLabel.getCurrentOrNullObject;
    expect(getCurrentLabel).toHaveBeenCalledTimes(1);
    expect(getCurrentLabel.mock.results[0].value.load).toHaveBeenCalledWith(
      'color,id,isEnabled,name,order,protectionType,siteId,tooltip'
    );
    expect(mockContext.sync).toHaveBeenCalledTimes(1);
    expect(consoleSpy).toHaveBeenCalledWith(isNullObject
      ? "The workbook doesn't have a sensitivity label."
      : expect.stringContaining('Current workbook sensitivity label:\n\nGeneral\n  ID: general-label'));
    expect(consoleErrorSpy).not.toHaveBeenCalled();
  });

  test.each(['Success', 'DocumentReadOnly', 'LabelNotFound', 'AbacAttributesRequired'])(
    'Excel: setting a workbook label handles result %s',
    async (updateResult) => {
      const { mockContext } = await runExcelSnippetTest({
        snippetPath: sensitivitySnippetPath,
        buttonId: 'set-workbook-label',
        elementValues: { 'label-id': 'engineering-label' },
        mockOptions: { sensitivityLabelUpdateResult: updateResult },
      });
      expect(mockContext.workbook.sensitivityLabel.tryToUpdate).toHaveBeenCalledWith('engineering-label');
      expect(mockContext.sync).toHaveBeenCalledTimes(1);
      if (updateResult === 'Success') {
        expect(consoleSpy).toHaveBeenCalledWith('Applied the sensitivity label to the workbook.');
        expect(consoleErrorSpy).not.toHaveBeenCalled();
      } else {
        expect(consoleErrorSpy).toHaveBeenCalledWith(`The sensitivity label wasn't applied. Result: ${updateResult}`);
        expect(consoleSpy).not.toHaveBeenCalled();
      }
    }
  );

  test('Excel: setting a label without a selection warns without updating the workbook', async () => {
    const warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
    const { mockContext } = await runExcelSnippetTest({
      snippetPath: sensitivitySnippetPath,
      buttonId: 'set-workbook-label',
    });
    expect(warnSpy).toHaveBeenCalledWith('Get available labels and select a label first.');
    expect(mockContext.workbook.sensitivityLabel.tryToUpdate).not.toHaveBeenCalled();
    expect(mockContext.sync).not.toHaveBeenCalled();
    expect(consoleErrorSpy).not.toHaveBeenCalled();
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
