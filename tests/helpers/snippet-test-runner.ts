import { TestSnippet, loadSnippetByPath } from './test-helpers';
import { executeSnippetCode, createMockDocument, clickButton } from './snippet-runtime';
import {
  createExcelMock,
  createWordMock,
  createPowerPointMock,
  createOneNoteMock,
  createOutlookMock,
  createOfficeCommonApiMock,
  ExcelMockOptions,
  WordMockOptions,
  PowerPointMockOptions,
  OneNoteMockOptions,
  OutlookMockOptions,
} from './mock-factories';

/**
 * Test context for snippet execution
 */
export interface SnippetTestContext {
  snippet: TestSnippet;
  buttonHandlers: Map<string, Function>;
  consoleSpy: jest.SpyInstance;
}

/**
 * Options for running a snippet test
 */
export interface RunSnippetTestOptions {
  snippetPath: string;
  buttonId?: string;
  mockOptions?: any;
  assertions?: (context: any) => void;
  skipButtonClick?: boolean;
}

/**
 * Create a test context with spies and handlers
 */
export function createTestContext(): SnippetTestContext {
  return {
    snippet: null as any,
    buttonHandlers: new Map(),
    consoleSpy: jest.spyOn(console, 'log').mockImplementation(),
  };
}

/**
 * Clean up test context
 */
export function cleanupTestContext(context: SnippetTestContext) {
  jest.restoreAllMocks();
}

/**
 * Run an Excel snippet test
 */
export async function runExcelSnippetTest(options: RunSnippetTestOptions) {
  const { snippetPath, buttonId = 'run', mockOptions = {}, assertions, skipButtonClick = false } = options;

  const snippet = loadSnippetByPath(snippetPath);
  const buttonHandlers = new Map<string, Function>();

  // Create Excel mock
  const { mockObject, mockContext, mockRange, mockWorksheet } = createExcelMock(mockOptions as ExcelMockOptions);

  (global as any).Excel = mockObject;
  (global as any).document = createMockDocument(buttonHandlers);

  // Execute snippet
  executeSnippetCode(snippet);

  // Trigger button click if needed
  if (!skipButtonClick) {
    await clickButton(buttonHandlers, buttonId);
  }

  // Run custom assertions
  if (assertions) {
    assertions({ mockContext, mockRange, mockWorksheet });
  }

  return { mockContext, mockRange, mockWorksheet };
}

/**
 * Run a Word snippet test
 */
export async function runWordSnippetTest(options: RunSnippetTestOptions) {
  const { snippetPath, buttonId = 'run', mockOptions = {}, assertions, skipButtonClick = false } = options;

  const snippet = loadSnippetByPath(snippetPath);
  const buttonHandlers = new Map<string, Function>();

  // Create Word mock
  const { mockObject, mockContext, mockRange, mockBody, mockParagraph } = createWordMock(
    mockOptions as WordMockOptions
  );

  (global as any).Word = mockObject;
  (global as any).document = createMockDocument(buttonHandlers);

  // Execute snippet
  executeSnippetCode(snippet);

  // Trigger button click if needed
  if (!skipButtonClick) {
    await clickButton(buttonHandlers, buttonId);
  }

  // Run custom assertions
  if (assertions) {
    assertions({ mockContext, mockRange, mockBody, mockParagraph });
  }

  return { mockContext, mockRange, mockBody, mockParagraph };
}

/**
 * Run a PowerPoint snippet test
 */
export async function runPowerPointSnippetTest(options: RunSnippetTestOptions) {
  const { snippetPath, buttonId = 'run', mockOptions = {}, assertions, skipButtonClick = false } = options;

  const snippet = loadSnippetByPath(snippetPath);
  const buttonHandlers = new Map<string, Function>();

  // Create PowerPoint mock
  const { mockObject, mockContext, mockSlides, mockSlide, mockShapes, mockShape } = createPowerPointMock(
    mockOptions as PowerPointMockOptions
  );

  (global as any).PowerPoint = mockObject;
  (global as any).document = createMockDocument(buttonHandlers);

  // Execute snippet
  executeSnippetCode(snippet);

  // Trigger button click if needed
  if (!skipButtonClick) {
    await clickButton(buttonHandlers, buttonId);
  }

  // Run custom assertions
  if (assertions) {
    assertions({ mockContext, mockSlides, mockSlide, mockShapes, mockShape });
  }

  return { mockContext, mockSlides, mockSlide, mockShapes, mockShape };
}

/**
 * Run a OneNote snippet test
 */
export async function runOneNoteSnippetTest(options: RunSnippetTestOptions) {
  const { snippetPath, buttonId = 'run', mockOptions = {}, assertions, skipButtonClick = false } = options;

  const snippet = loadSnippetByPath(snippetPath);
  const buttonHandlers = new Map<string, Function>();

  // Create OneNote mock
  const { mockObject, mockContext, mockPage, mockSection, mockOutline } = createOneNoteMock(
    mockOptions as OneNoteMockOptions
  );

  (global as any).OneNote = mockObject;
  (global as any).document = createMockDocument(buttonHandlers);

  // Execute snippet
  executeSnippetCode(snippet);

  // Trigger button click if needed
  if (!skipButtonClick) {
    await clickButton(buttonHandlers, buttonId);
  }

  // Run custom assertions
  if (assertions) {
    assertions({ mockContext, mockPage, mockSection, mockOutline });
  }

  return { mockContext, mockPage, mockSection, mockOutline };
}

/**
 * Run an Outlook snippet test
 */
export async function runOutlookSnippetTest(options: RunSnippetTestOptions) {
  const { snippetPath, buttonId = 'run', mockOptions = {}, assertions, skipButtonClick = false } = options;

  const snippet = loadSnippetByPath(snippetPath);
  const buttonHandlers = new Map<string, Function>();

  // Create Outlook mock
  const { Office: OutlookOffice, mockMailbox, mockItem } = createOutlookMock(mockOptions as OutlookMockOptions);

  (global as any).Office = OutlookOffice;
  (global as any).document = createMockDocument(buttonHandlers);

  // Execute snippet
  executeSnippetCode(snippet);

  // Trigger button click if needed
  if (!skipButtonClick) {
    await clickButton(buttonHandlers, buttonId);
  }

  // Run custom assertions
  if (assertions) {
    assertions({ mockMailbox, mockItem });
  }

  return { mockMailbox, mockItem };
}

/**
 * Run a Common API snippet test
 */
export async function runCommonApiSnippetTest(options: RunSnippetTestOptions) {
  const { snippetPath, buttonId = 'run', assertions, skipButtonClick = false } = options;

  const snippet = loadSnippetByPath(snippetPath);
  const buttonHandlers = new Map<string, Function>();

  // Create Common API mock
  (global as any).Office = createOfficeCommonApiMock();
  (global as any).document = createMockDocument(buttonHandlers);

  // Execute snippet
  executeSnippetCode(snippet);

  // Trigger button click if needed
  if (!skipButtonClick) {
    await clickButton(buttonHandlers, buttonId);
  }

  // Run custom assertions
  if (assertions) {
    const Office = (global as any).Office;
    assertions({ Office });
  }

  return { Office: (global as any).Office };
}
