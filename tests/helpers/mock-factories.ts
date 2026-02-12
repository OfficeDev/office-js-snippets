import { OfficeMockObject } from 'office-addin-mock';
import { OfficeApp } from 'office-addin-manifest';

/**
 * Factory for creating Office API mocks
 *
 * IMPORTANT LIMITATIONS:
 *
 * These mocks provide basic API structure for syntax testing but do NOT replicate
 * real Office.js behavior. They are suitable for smoke tests only.
 *
 * What these mocks provide:
 * ✅ Basic object structure (context, workbook, worksheet, range, etc.)
 * ✅ Common properties (address, text, name, etc.)
 * ✅ Method stubs (load, sync, getRange, etc.)
 * ✅ Simple return values (fixed strings, numbers)
 *
 * What these mocks do NOT provide:
 * ❌ Dynamic collections - items[] arrays never change
 * ❌ load/sync semantics - properties are always available immediately
 * ❌ State updates - mutations don't affect subsequent reads
 * ❌ Real batching - sync() doesn't actually batch requests
 * ❌ Error simulation - mocks don't throw Office.js errors
 * ❌ Ordering guarantees - collection item order may not match real Office
 *
 * Examples of what won't work correctly:
 *
 * // Collections don't update:
 * body.insertParagraph("New text", "End");
 * await context.sync();
 * console.log(paragraphs.items.length); // Still returns original length!
 *
 * // load() doesn't control availability:
 * range.load("address"); // Does nothing
 * await context.sync();   // Does nothing
 * console.log(range.address); // Works anyway (shouldn't without load)
 *
 * // Iterations use static data:
 * paragraphs.items.forEach(p => {
 *   console.log(p.text); // Always same paragraphs, never changes
 * });
 *
 * Use these mocks for:
 * ✅ Verifying code syntax is correct
 * ✅ Checking API method names aren't typos
 * ✅ Basic smoke testing that code runs
 *
 * Do NOT rely on these mocks for:
 * ❌ Verifying correct Office.js behavior
 * ❌ Testing collection operations
 * ❌ Validating load/sync patterns
 * ❌ Confirming output correctness
 */

// Common Office API mocks
export function createOfficeCommonApiMock() {
  return {
    AsyncResultStatus: {
      Succeeded: 'succeeded',
      Failed: 'failed',
    },
    CoercionType: {
      Text: 'text',
      Matrix: 'matrix',
      Table: 'table',
      Html: 'html',
      Ooxml: 'ooxml',
    },
    context: {
      document: {
        getSelectedDataAsync: jest.fn(
          (coercionType: any, optionsOrCallback?: any, callback?: Function) => {
            // Handle both 2-arg (coercionType, callback) and 3-arg (coercionType, options, callback) overloads
            const cb: Function =
              typeof optionsOrCallback === 'function' ? optionsOrCallback : (callback as Function);
            cb({ status: 'succeeded', value: 'Mock selected data' });
          }
        ),
        setSelectedDataAsync: jest.fn(
          (data: any, optionsOrCallback?: any, callback?: Function) => {
            // Handle both 2-arg (data, callback) and 3-arg (data, options, callback) overloads
            const cb: Function =
              typeof optionsOrCallback === 'function' ? optionsOrCallback : (callback as Function);
            cb({ status: 'succeeded' });
          }
        ),
      },
    },
  };
}

// Excel-specific mocks
export interface ExcelMockOptions {
  rangeAddress?: string;
  rangeValues?: any[][];
  rangeText?: string;
  worksheetName?: string;
}

export function createExcelMock(options: ExcelMockOptions = {}) {
  const {
    rangeAddress = 'A1:B2',
    rangeValues = [['value1', 'value2']],
    rangeText = 'Mock text',
    worksheetName = 'Sheet1',
  } = options;

  const mockRange = {
    address: rangeAddress,
    values: rangeValues,
    text: rangeText,
    format: {
      fill: { color: '' },
      font: { color: '', bold: false, italic: false },
    },
    load: jest.fn(),
  };

  const mockWorksheet = {
    name: worksheetName,
    getRange: jest.fn(() => mockRange),
    getRangeByIndexes: jest.fn(() => mockRange),
  };

  const mockContext = {
    workbook: {
      getSelectedRange: jest.fn(() => mockRange),
      worksheets: {
        getActiveWorksheet: jest.fn(() => mockWorksheet),
        getItem: jest.fn(() => mockWorksheet),
        add: jest.fn(() => mockWorksheet),
      },
    },
    sync: jest.fn().mockResolvedValue(undefined),
  };

  const mockData = {
    run: jest.fn(async (callback: Function) => {
      await callback(mockContext);
    }),
  };

  return {
    mockObject: new OfficeMockObject(mockData, OfficeApp.Excel),
    mockContext,
    mockRange,
    mockWorksheet,
  };
}

// Word-specific mocks
export interface WordMockOptions {
  selectionText?: string;
  paragraphText?: string;
  bodyText?: string;
}

export function createWordMock(options: WordMockOptions = {}) {
  const {
    selectionText = 'Mock selected text',
    paragraphText = 'Mock paragraph',
    bodyText = 'Mock body text',
  } = options;

  // Mock for getText() return value
  const mockTextValue = {
    value: paragraphText,
  };

  const mockParagraph = {
    text: paragraphText,
    font: { color: '', bold: false },
    load: jest.fn(),
    getText: jest.fn(() => mockTextValue),
  };

  // Mock paragraph collection
  const mockParagraphs = {
    getFirst: jest.fn(() => mockParagraph),
    items: [mockParagraph],
  };

  // Mock table cell body
  const mockCellBody = {
    text: 'Mock cell text',
    load: jest.fn(),
  };

  // Mock table
  const mockTable = {
    getCell: jest.fn(() => ({ body: mockCellBody })),
    load: jest.fn(),
  };

  // Mock table collection
  const mockTables = {
    getFirst: jest.fn(() => mockTable),
    items: [mockTable],
  };

  const mockRange = {
    text: selectionText,
    font: { color: '', bold: false },
    load: jest.fn(),
    insertText: jest.fn(),
    insertParagraph: jest.fn(() => mockParagraph),
    paragraphs: mockParagraphs,
  };

  const mockBody = {
    text: bodyText,
    insertParagraph: jest.fn(() => mockParagraph),
    insertText: jest.fn(),
    insertTable: jest.fn(() => mockTable),
    load: jest.fn(),
    tables: mockTables,
  };

  // Mock document properties
  const mockProperties = {
    author: 'Mock Author',
    title: 'Mock Title',
    subject: 'Mock Subject',
    comments: 'Mock Comments',
    keywords: 'Mock Keywords',
    manager: 'Mock Manager',
    company: 'Mock Company',
    category: 'Mock Category',
    applicationName: 'Microsoft Word',
    creationDate: new Date('2024-01-01'),
    lastAuthor: 'Mock Last Author',
    lastPrintDate: new Date('2024-01-02'),
    lastSaveTime: new Date('2024-01-03'),
    revisionNumber: 1,
    template: 'Normal.dotm',
    load: jest.fn(),
  };

  const mockContext = {
    document: {
      body: mockBody,
      getSelection: jest.fn(() => mockRange),
      properties: mockProperties,
      compare: jest.fn(),
    },
    sync: jest.fn().mockResolvedValue(undefined),
  };

  const mockData = {
    run: jest.fn(async (callback: Function) => {
      await callback(mockContext);
    }),
  };

  // Create the mock object
  const wordMockObject = new OfficeMockObject(mockData, OfficeApp.Word) as any;

  // Add enums that some snippets need
  wordMockObject.CompareTarget = {
    compareTargetCurrent: 'Current',
    compareTargetNew: 'New',
  };

  return {
    mockObject: wordMockObject,
    mockContext,
    mockRange,
    mockBody,
    mockParagraph,
    mockParagraphs,
    mockTable,
    mockTables,
    mockProperties,
  };
}

// PowerPoint-specific mocks
export interface PowerPointMockOptions {
  slideCount?: number;
  shapeCount?: number;
}

export function createPowerPointMock(options: PowerPointMockOptions = {}) {
  const { slideCount = 1, shapeCount = 0 } = options;

  const mockShape = {
    id: 'shape-1',
    name: 'Shape 1',
    textFrame: {
      textRange: {
        text: 'Mock text',
      },
    },
  };

  const mockShapes = {
    items: Array(shapeCount).fill(mockShape),
    addTextBox: jest.fn(() => mockShape),
    addShape: jest.fn(() => mockShape),
    load: jest.fn(),
  };

  const mockSlide = {
    id: 'slide-1',
    shapes: mockShapes,
    load: jest.fn(),
  };

  const mockSlides = {
    items: Array(slideCount).fill(mockSlide),
    getItemAt: jest.fn((index: number) => mockSlide),
    add: jest.fn(() => mockSlide),
    load: jest.fn(),
  };

  const mockContext = {
    presentation: {
      slides: mockSlides,
      load: jest.fn(),
    },
    sync: jest.fn().mockResolvedValue(undefined),
  };

  const mockData = {
    run: jest.fn(async (callback: Function) => {
      await callback(mockContext);
    }),
  };

  return {
    mockObject: new OfficeMockObject(mockData, OfficeApp.PowerPoint),
    mockContext,
    mockSlides,
    mockSlide,
    mockShapes,
    mockShape,
  };
}

// OneNote-specific mocks
export interface OneNoteMockOptions {
  pageTitle?: string;
  outlineText?: string;
}

export function createOneNoteMock(options: OneNoteMockOptions = {}) {
  const { pageTitle = 'Mock Page' } = options;

  const mockOutline = {
    id: 'outline-1',
    appendHtml: jest.fn(),
    appendRichText: jest.fn(),
  };

  const mockPage = {
    title: pageTitle,
    contents: {
      items: [mockOutline],
    },
    addOutline: jest.fn(() => mockOutline),
    load: jest.fn(),
  };

  const mockSection = {
    name: 'Mock Section',
    addPage: jest.fn(() => mockPage),
    getActivePageOrNull: jest.fn(() => mockPage),
  };

  const mockContext = {
    application: {
      getActiveSection: jest.fn(() => mockSection),
      getActivePage: jest.fn(() => mockPage),
    },
    sync: jest.fn().mockResolvedValue(undefined),
  };

  const mockData = {
    run: jest.fn(async (callback: Function) => {
      await callback(mockContext);
    }),
  };

  return {
    mockObject: new OfficeMockObject(mockData, OfficeApp.OneNote),
    mockContext,
    mockPage,
    mockSection,
    mockOutline,
  };
}

// Outlook-specific mocks (Outlook has a different API structure)
export interface OutlookMockOptions {
  subject?: string;
  body?: string;
  from?: string;
}

export function createOutlookMock(options: OutlookMockOptions = {}) {
  const {
    subject = 'Mock Email Subject',
    body = 'Mock email body',
    from = 'sender@example.com',
  } = options;

  const mockItem = {
    subject: subject,
    body: {
      getAsync: jest.fn((callback: Function) => {
        callback({ value: body, status: 'succeeded' });
      }),
      setAsync: jest.fn((data: any, options: any, callback: Function) => {
        if (callback) callback({ status: 'succeeded' });
      }),
    },
    from: {
      emailAddress: from,
    },
    to: {
      getAsync: jest.fn((callback: Function) => {
        callback({ value: [], status: 'succeeded' });
      }),
    },
  };

  const mockMailbox = {
    item: mockItem,
    userProfile: {
      emailAddress: 'user@example.com',
      displayName: 'Test User',
    },
  };

  return {
    Office: {
      context: {
        mailbox: mockMailbox,
      },
    },
    mockMailbox,
    mockItem,
  };
}

// Project-specific mocks
export function createProjectMock() {
  const mockTask = {
    name: 'Mock Task',
    id: 'task-1',
  };

  const mockResource = {
    name: 'Mock Resource',
    id: 'resource-1',
  };

  return {
    Office: {
      context: {
        document: {
          getSelectedTaskAsync: jest.fn((callback: Function) => {
            callback({ value: mockTask, status: 'succeeded' });
          }),
          getSelectedResourceAsync: jest.fn((callback: Function) => {
            callback({ value: mockResource, status: 'succeeded' });
          }),
        },
      },
    },
    mockTask,
    mockResource,
  };
}
