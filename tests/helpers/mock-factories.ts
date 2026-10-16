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
    FileType: {
      Pdf: 'pdf',
      Text: 'text',
      Compressed: 'compressed',
    },
    context: {
      requirements: {
        isSetSupported: jest.fn(() => true),
      },
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
        // getFileAsync is a no-op stub; callback is intentionally never invoked so that
        // browser-only code in the callback (Blob, window, URL.createObjectURL) is not reached
        // in the Node.js test environment.
        getFileAsync: jest.fn(),
      },
    },
  };
}

// Excel-specific mocks
export interface ExcelMockOptions {
  rangeAddress?: string;
  rangeValues?: any[][];
  rangeValuesAsJson?: unknown[][];
  rangeText?: string;
  worksheetName?: string;
  horizontalPageBreakCount?: number;
}

export function createExcelMock(options: ExcelMockOptions = {}) {
  const {
    rangeAddress = 'A1:B2',
    rangeValues = [['value1', 'value2']],
    rangeValuesAsJson = [[{ type: 'String', basicValue: 'Mock value' }]],
    rangeText = 'Mock text',
    worksheetName = 'Sheet1',
    horizontalPageBreakCount = 2,
  } = options;

  const mockFont = {
    color: '',
    bold: false,
    italic: false,
    name: 'Calibri',
    size: 11,
    underline: 'None',
    load: jest.fn(),
    set: jest.fn((properties: Record<string, unknown>) => Object.assign(mockFont, properties)),
  };

  const mockRange = {
    address: rangeAddress,
    values: rangeValues,
    valuesAsJson: rangeValuesAsJson,
    text: rangeText,
    numberFormat: [],
    format: {
      fill: { color: '' },
      font: mockFont,
      autofitColumns: jest.fn(),
      columnWidth: 0,
      rowHeight: 0,
    },
    load: jest.fn(),
  };

  const mockTableRow = {
    index: 0,
    values: [['Pau Pons', 'North', 1400]],
    load: jest.fn(),
  };

  const mockTableRows = {
    items: [mockTableRow],
    count: 1,
    add: jest.fn(() => mockTableRow),
    deleteRows: jest.fn(),
    getCount: jest.fn(() => ({ value: 1 })),
    getItemAt: jest.fn(() => mockTableRow),
    load: jest.fn(),
  };

  const mockTable = {
    name: 'SalesTable',
    rows: mockTableRows,
  };

  function createPageBreakCollection(count: number, horizontal: boolean) {
    const pageBreaks = Array.from({ length: count }, (_, index) => ({
      rowIndex: horizontal ? 11 + index * 10 : 0,
      columnIndex: horizontal ? 0 : 4,
      load: jest.fn(),
      delete: jest.fn(),
      getCellAfterBreak: jest.fn(() => mockRange),
    }));
    return {
      items: pageBreaks,
      load: jest.fn(),
      add: jest.fn(() => ({
        load: jest.fn(),
        delete: jest.fn(),
        getCellAfterBreak: jest.fn(() => mockRange),
      })),
      getCount: jest.fn(() => ({ value: pageBreaks.length })),
      getItem: jest.fn((index: number) => {
        if (index < 0 || index >= pageBreaks.length) {
          throw new RangeError('Page break index is out of range.');
        }
        return pageBreaks[index];
      }),
      removePageBreaks: jest.fn(),
    };
  }

  const mockWorksheet = {
    name: worksheetName,
    isNullObject: false,
    horizontalPageBreaks: createPageBreakCollection(horizontalPageBreakCount, true),
    verticalPageBreaks: createPageBreakCollection(1, false),
    pageLayout: {
      setPrintArea: jest.fn(),
      zoom: { scale: 100 },
    },
    getRange: jest.fn(() => mockRange),
    getRangeByIndexes: jest.fn(() => mockRange),
    getUsedRange: jest.fn(() => mockRange),
    tables: {
      add: jest.fn(() => mockTable),
      getItem: jest.fn(() => mockTable),
    },
    load: jest.fn(),
    delete: jest.fn(),
    activate: jest.fn(),
  };

  const mockContext = {
    workbook: {
      getSelectedRange: jest.fn(() => mockRange),
      worksheets: {
        getActiveWorksheet: jest.fn(() => mockWorksheet),
        getItem: jest.fn(() => mockWorksheet),
        getItemOrNullObject: jest.fn(() => mockWorksheet),
        add: jest.fn(() => mockWorksheet),
      },
    },
    sync: jest.fn().mockResolvedValue(undefined),
  };

  const mockData = {
    RangeUnderlineStyle: {
      none: 'None',
      single: 'Single',
      double: 'Double',
      singleAccountant: 'SingleAccountant',
      doubleAccountant: 'DoubleAccountant',
    },
    run: jest.fn(async (callback: Function) => {
      await callback(mockContext);
    }),
  };

  const mockObject = new OfficeMockObject(mockData, OfficeApp.Excel);
  Object.assign(mockObject, {
    CellValueType: {
      double: 'Double',
      entity: 'Entity',
      error: 'Error',
      function: 'Function',
      notAvailable: 'NotAvailable',
    },
    ErrorCellValueType: {
      busy: 'Busy',
    },
    FunctionCellValueType: {
      javaScriptReference: 'JavaScriptReference',
    },
  });

  return {
    mockObject,
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

  const mockBorder = {
    isVisible: false,
    color: '',
    lineStyle: '',
    lineWidth: '',
    load: jest.fn(),
  };

  const mockBorders = {
    items: [mockBorder],
    outsideBorderColor: '#C65911',
    outsideBorderType: 'Single',
    outsideBorderWidth: 'Pt100',
    load: jest.fn(),
  };

  const mockThreeDimensionalFormat = {
    isVisible: false,
    bevelTopType: 'none',
    bevelTopDepth: 0,
    bevelTopInset: 0,
    load: jest.fn(),
  };

  const mockParagraph = {
    text: paragraphText,
    alignment: 'Left',
    firstLineIndent: 0,
    leftIndent: 0,
    rightIndent: 0,
    lineSpacing: 12,
    spaceBefore: 0,
    spaceAfter: 0,
    lineUnitBefore: 0,
    lineUnitAfter: 0,
    style: '',
    font: {
      color: '',
      bold: false,
      threeDimensionalFormat: mockThreeDimensionalFormat,
    },
    borders: mockBorders,
    load: jest.fn(),
    getText: jest.fn(() => mockTextValue),
    getRange: jest.fn(),
  };

  // Mock paragraph collection
  const mockParagraphs = {
    getFirst: jest.fn(() => mockParagraph),
    getLast: jest.fn(() => mockParagraph),
    items: [mockParagraph],
    load: jest.fn(),
  };

  const mockParagraphFormat = {
    alignment: 'Left',
    firstLineIndent: 0,
    keepTogether: false,
    keepWithNext: false,
    leftIndent: 0,
    lineSpacing: 12,
    rightIndent: 0,
    spaceAfter: 8,
    spaceBefore: 0,
    widowControl: true,
    load: jest.fn(),
    toJSON: jest.fn(function() {
      return {
        alignment: this.alignment,
        firstLineIndent: this.firstLineIndent,
        keepTogether: this.keepTogether,
        keepWithNext: this.keepWithNext,
        leftIndent: this.leftIndent,
        lineSpacing: this.lineSpacing,
        rightIndent: this.rightIndent,
        spaceAfter: this.spaceAfter,
        spaceBefore: this.spaceBefore,
        widowControl: this.widowControl,
      };
    }),
  };

  const mockShading = {
    backgroundPatternColor: '#FFFFFF',
    foregroundPatternColor: '#FFFFFF',
    texture: 'None',
    load: jest.fn(),
  };

  const mockStyleFont = {
    color: '#000000',
    name: 'Calibri',
    size: 11,
    load: jest.fn(),
  };

  const mockStyle = {
    isNullObject: false,
    name: 'SampleParagraphFormat',
    nameLocal: 'Sample Style',
    type: 'Paragraph',
    borders: mockBorders,
    font: mockStyleFont,
    paragraphFormat: mockParagraphFormat,
    shading: mockShading,
    delete: jest.fn(),
    load: jest.fn(),
  };

  const mockStyles = {
    getByName: jest.fn(() => mockStyle),
    getByNameOrNullObject: jest.fn(() => mockStyle),
    load: jest.fn(),
  };

  // Mock table cell body
  const mockCellBody = {
    font: { color: '' },
    text: 'Mock cell text',
    load: jest.fn(),
  };

  const mockTableCell = {
    body: mockCellBody,
    cellIndex: 0,
    columnWidth: 72,
    horizontalAlignment: 'Left',
    rowIndex: 1,
    shadingColor: '#FFFFFF',
    value: 'North',
    verticalAlignment: 'Top',
    width: 72,
    load: jest.fn(),
  };

  const mockTableColumn = {
    columnIndex: 0,
    isFirst: true,
    isLast: false,
    width: 72,
    load: jest.fn(),
  };

  const mockTableColumns = {
    items: [mockTableColumn],
    load: jest.fn(),
  };

  const mockTableRowCollection: any = {
    items: [],
    getFirst: jest.fn(),
    load: jest.fn(),
  };

  const mockTableRow: any = {
    cellCount: 3,
    font: { color: '' },
    horizontalAlignment: 'Left',
    isHeader: false,
    preferredHeight: 18,
    rowIndex: 1,
    shadingColor: '#FFFFFF',
    values: [['North', '$125,000', 'On track']],
    verticalAlignment: 'Top',
    delete: jest.fn(),
    getNext: jest.fn(),
    insertRows: jest.fn(() => mockTableRowCollection),
    load: jest.fn(),
  };
  mockTableRow.getNext.mockReturnValue(mockTableRow);
  mockTableRowCollection.items = [mockTableRow];
  mockTableRowCollection.getFirst.mockReturnValue(mockTableRow);

  // Mock table
  const mockTable = {
    getCell: jest.fn(() => mockTableCell),
    getRange: jest.fn(),
    rows: mockTableRowCollection,
    load: jest.fn(),
  };

  // Mock table collection
  const mockTables = {
    getFirst: jest.fn(() => mockTable),
    items: [mockTable],
    load: jest.fn(),
  };

  const mockRange = {
    text: selectionText,
    font: { color: '', bold: false, highlightColor: '', underline: 'None' },
    delete: jest.fn(),
    load: jest.fn(),
    insertText: jest.fn(),
    insertParagraph: jest.fn(() => mockParagraph),
    paragraphs: mockParagraphs,
    tableColumns: mockTableColumns,
  };
  mockTable.getRange.mockReturnValue(mockRange);
  mockParagraph.getRange.mockReturnValue(mockRange);

  const mockBreak = {
    pageIndex: 1,
    range: mockRange,
    load: jest.fn(),
  };

  const mockBreaks = {
    items: [mockBreak],
    load: jest.fn(),
  };

  const mockPage = {
    index: 1,
    height: 792,
    width: 612,
    isNullObject: false,
    breaks: mockBreaks,
    getNext: jest.fn(),
    getNextOrNullObject: jest.fn(),
    getRange: jest.fn(() => mockRange),
    load: jest.fn(),
    toJSON: jest.fn(function() {
      return { index: this.index, height: this.height, width: this.width };
    }),
  };

  // Fixed pages exercise navigation without modeling Word's pagination.
  const mockNullPage = { isNullObject: true };
  const mockNextPage = {
    ...mockPage,
    index: 2,
    breaks: { items: [], load: jest.fn() },
    getNext: jest.fn(() => { throw new Error('ItemNotFound'); }),
    getNextOrNullObject: jest.fn(() => mockNullPage),
    load: jest.fn(),
  };
  mockPage.getNext.mockReturnValue(mockNextPage);
  mockPage.getNextOrNullObject.mockReturnValue(mockNextPage);

  const mockPages = {
    items: [mockPage, mockNextPage],
    getFirst: jest.fn(() => mockPage),
    getFirstOrNullObject: jest.fn(() => mockPage),
    load: jest.fn(),
  };
  Object.assign(mockRange, { pages: mockPages });

  const mockPane = {
    pages: mockPages,
    pagesEnclosingViewport: mockPages,
    load: jest.fn(),
  };

  const mockWindow = {
    activePane: mockPane,
    panes: {
      items: [mockPane],
      load: jest.fn(),
    },
    load: jest.fn(),
  };

  const mockTableOfAuthorities = {
    category: 1,
    entrySeparator: ' ',
    isCategoryHeaderIncluded: true,
    isPassimUsed: false,
    range: mockRange,
    delete: jest.fn(),
    load: jest.fn(),
  };

  const mockTableOfContents = {
    areBuiltInHeadingStylesUsed: true,
    arePageNumbersIncluded: true,
    arePageNumbersRightAligned: true,
    lowerHeadingLevel: 2,
    range: mockRange,
    tabLeader: 'Dots',
    upperHeadingLevel: 1,
    delete: jest.fn(),
    load: jest.fn(),
    updatePageNumbers: jest.fn(),
  };

  const mockTableOfFigures = {
    areFieldsUsed: true,
    arePageNumbersIncluded: true,
    arePageNumbersRightAligned: true,
    captionLabel: 'Figure',
    isLabelIncluded: true,
    range: mockRange,
    tabLeader: 'Dots',
    delete: jest.fn(),
    load: jest.fn(),
    updatePageNumbers: jest.fn(),
  };

  const mockTablesOfAuthorities = {
    items: [mockTableOfAuthorities],
    add: jest.fn(() => mockTableOfAuthorities),
    markAllCitations: jest.fn(),
    load: jest.fn(),
  };

  const mockTablesOfContents = {
    items: [mockTableOfContents],
    add: jest.fn(() => mockTableOfContents),
    load: jest.fn(),
  };

  const mockTablesOfFigures = {
    items: [mockTableOfFigures],
    add: jest.fn(() => mockTableOfFigures),
    markTocEntry: jest.fn(),
    load: jest.fn(),
  };

  const mockTrackedChange = {
    author: 'Mock Author',
    date: new Date('2024-01-01'),
    isNullObject: false,
    text: 'Mock tracked change',
    type: 'Added',
    accept: jest.fn(),
    getNext: jest.fn(),
    getRange: jest.fn(() => mockRange),
    load: jest.fn(),
    reject: jest.fn(),
  };
  mockTrackedChange.getNext.mockReturnValue(mockTrackedChange);

  const mockRevision = {
    author: 'Mock Author',
    date: new Date('2024-01-01'),
    formatDescription: '',
    index: 1,
    type: 'Insert',
    get range() { return mockRange; },
    get movedRange() { return mockRange; },
    accept: jest.fn(),
    reject: jest.fn(),
  };

  const mockTrackedChanges = {
    items: [mockTrackedChange],
    acceptAll: jest.fn(),
    getFirst: jest.fn(() => mockTrackedChange),
    getFirstOrNullObject: jest.fn(() => mockTrackedChange),
    load: jest.fn(),
    rejectAll: jest.fn(),
  };

  const mockBody = {
    text: bodyText,
    paragraphs: mockParagraphs,
    clear: jest.fn(),
    insertParagraph: jest.fn(() => mockParagraph),
    insertBreak: jest.fn(),
    insertText: jest.fn(),
    insertTable: jest.fn(() => mockTable),
    getRange: jest.fn(() => mockRange),
    getTrackedChanges: jest.fn(() => mockTrackedChanges),
    load: jest.fn(),
    search: jest.fn(() => ({
      getFirst: jest.fn(() => mockRange),
    })),
    tables: mockTables,
  };

  const mockSectionBody = {
    ...mockBody,
    text: 'Mock section body',
    load: jest.fn(),
  };
  const mockNextSectionBody = {
    ...mockBody,
    text: 'Mock next section body',
    load: jest.fn(),
  };
  const mockPageSetup = {
    sectionStart: 'NewPage',
    load: jest.fn(),
    set: jest.fn(),
    toJSON: jest.fn(() => ({ sectionStart: 'NewPage' })),
  };
  const mockSectionBorder = {
    color: '#000000',
    isVisible: false,
    lineStyle: 'Single',
    lineWidth: 'Pt050',
    load: jest.fn(),
  };
  const mockSectionBorders = {
    items: [mockSectionBorder],
    applyPageBordersToAllSections: jest.fn(),
    load: jest.fn(),
  };
  const mockNextSection = {
    body: mockNextSectionBody,
    borders: mockSectionBorders,
    isNullObject: false,
    pageSetup: { ...mockPageSetup },
    protectedForForms: false,
    load: jest.fn(),
  };
  const mockSection = {
    body: mockSectionBody,
    borders: mockSectionBorders,
    isNullObject: false,
    pageSetup: { ...mockPageSetup },
    protectedForForms: false,
    getNext: jest.fn(() => mockNextSection),
    getNextOrNullObject: jest.fn(() => mockNextSection),
    load: jest.fn(),
  };
  Object.assign(mockRange, { parentBody: mockSectionBody });
  Object.assign(mockSectionBody, { parentSection: mockSection });
  const mockThirdSection = {
    ...mockNextSection,
    pageSetup: { ...mockPageSetup },
  };
  const mockSections = {
    items: [mockSection, mockNextSection, mockThirdSection],
    getFirst: jest.fn(() => mockSection),
    load: jest.fn(),
  };

  // Mock coauthoring objects
  const mockCoauthor = {
    name: 'Mock Author',
    emailAddress: 'author@example.com',
    isMe: false,
    id: 'author-1',
    locks: {
      items: [] as any[],
      load: jest.fn(),
    },
    load: jest.fn(),
  };

  const mockMe = {
    name: 'Mock Me',
    emailAddress: 'me@example.com',
    id: 'me-1',
    load: jest.fn(),
  };

  const mockCoauthoringLock = {
    type: 'reservation',
    unlock: jest.fn(),
    load: jest.fn(),
  };

  const mockCoauthoringLockCollection = {
    items: [mockCoauthoringLock],
    add: jest.fn(() => mockCoauthoringLock),
    unlockEphemeralLocks: jest.fn(),
    load: jest.fn(),
  };

  const mockConflict = {
    type: 'text',
    accept: jest.fn(),
    reject: jest.fn(),
    load: jest.fn(),
  };

  const mockConflictCollection = {
    items: [mockConflict],
    acceptAll: jest.fn(),
    rejectAll: jest.fn(),
    load: jest.fn(),
  };

  const mockCoauthoringUpdate = {
    range: {
      text: 'Mock update text',
      load: jest.fn(),
    },
    load: jest.fn(),
  };

  const mockUpdateCollection = {
    items: [mockCoauthoringUpdate],
    load: jest.fn(),
  };

  const mockCoauthoring = {
    canCoauthor: true,
    canMerge: true,
    pendingUpdates: false,
    authors: {
      items: [mockCoauthor],
      load: jest.fn(),
    },
    me: mockMe,
    locks: mockCoauthoringLockCollection,
    conflicts: mockConflictCollection,
    updates: mockUpdateCollection,
    load: jest.fn(),
  };

  const mockNullRange = { isNullObject: true, load: jest.fn() };
  const mockRanges = {
    items: [mockRange],
    getFirst: jest.fn(() => mockRange),
    getLast: jest.fn(() => mockRange),
    load: jest.fn(),
  };
  const mockContentControl = { tag: 'RangeSampleSummary', title: 'Project summary', text: 'Project summary' };
  const mockInlinePicture = { width: 90, height: 39 };
  const mockList = { id: 1, insertParagraph: jest.fn(() => mockParagraph) };
  const mockShape = { name: 'RangeSamplePicture', type: 'Picture', width: 90, height: 39 };
  const mockEditor = { name: 'Everyone', id: 'Everyone', load: jest.fn() };
  const mockEditors = {
    addByType: jest.fn(() => mockEditor),
    getCount: jest.fn(() => ({ value: 1 })),
    getItemAt: jest.fn(() => mockEditor),
  };
  const mockFrame = { width: 200, height: 50, range: mockRange };
  const mockFrames = { items: [mockFrame], add: jest.fn(() => mockFrame), load: jest.fn() };
  const mockNote = { body: { text: 'Sample endnote' } };
  const mockEndnotes = { items: [mockNote], load: jest.fn() };
  const extendedRange = Object.assign(mockRange, {
    isNullObject: false,
    isEmpty: false,
    style: 'Normal',
    start: 0,
    end: 40,
    id: 'range-1',
    storyLength: 80,
    storyType: 'MainText',
    isEndOfRowMark: false,
    isTextVisibleOnScreen: true,
    bold: true,
    italic: true,
    boldBidirectional: false,
    italicBidirectional: false,
    case: 'TitleSentence',
    characterWidth: 'Half',
    combineCharacters: false,
    disableCharacterSpaceGrid: false,
    emphasisMark: 'None',
    fitTextWidth: 100,
    highlightColorIndex: 'Yellow',
    horizontalInVertical: 'None',
    kana: 'Katakana',
    twoLinesInOne: 'None',
    showAll: false,
    languageId: 'EnglishUS',
    languageIdFarEast: 'Japanese',
    languageIdOther: 'EnglishUS',
    languageDetected: false,
    hasNoProofing: false,
    spellingChecked: false,
    grammarChecked: false,
    underline: 'Single',
    hyperlink: 'https://learn.microsoft.com/office/dev/add-ins/word/',
    borders: mockBorders,
    shading: mockShading,
    contentControls: { items: [mockContentControl], load: jest.fn() },
    fields: { items: [{ code: 'QUOTE "Ready for review."', result: mockRange }], load: jest.fn() },
    inlinePictures: { items: [mockInlinePicture], load: jest.fn() },
    lists: { items: [mockList], load: jest.fn() },
    tables: mockTables,
    parentTable: mockTable,
    parentTableCell: mockTableCell,
    parentTableOrNullObject: { isNullObject: true, load: jest.fn() },
    parentTableCellOrNullObject: { isNullObject: true, load: jest.fn() },
    bookmarks: { items: [{ name: 'RangeSampleSummary', isEmpty: false, range: mockRange }], load: jest.fn() },
    hyperlinks: { items: [{ address: 'https://learn.microsoft.com/office/dev/add-ins/word/', range: mockRange }], load: jest.fn() },
    shapes: { items: [mockShape], load: jest.fn() },
    sections: mockSections,
    frames: mockFrames,
    editors: mockEditors,
    endnotes: mockEndnotes,
    conflicts: mockConflictCollection,
    revisions: { items: [mockRevision], load: jest.fn() },
    clear: jest.fn(),
    getRange: jest.fn(() => mockRange),
    compareLocationWith: jest.fn(() => ({ value: 'Before' })),
    expandTo: jest.fn(() => mockRange),
    expandToOrNullObject: jest.fn(() => mockRange),
    intersectWith: jest.fn(() => mockRange),
    intersectWithOrNullObject: jest.fn(() => mockRange),
    getTextRanges: jest.fn(() => mockRanges),
    getNextTextRange: jest.fn(() => mockRange),
    getNextTextRangeOrNullObject: jest.fn(() => mockNullRange),
    split: jest.fn(() => mockRanges),
    search: jest.fn(() => mockRanges),
    select: jest.fn(),
    highlight: jest.fn(),
    removeHighlight: jest.fn(),
    getHtml: jest.fn(() => ({ value: '<p>Mock paragraph</p>' })),
    getOoxml: jest.fn(() => ({ value: '<w:p>Mock paragraph</w:p>' })),
    insertHtml: jest.fn(() => mockRange),
    insertOoxml: jest.fn(() => mockRange),
    insertFileFromBase64: jest.fn(() => mockRange),
    insertBreak: jest.fn(),
    insertTable: jest.fn(() => mockTable),
    insertBookmark: jest.fn(),
    getBookmarks: jest.fn(() => ({ value: ['RangeSampleSummary'] })),
    getHyperlinkRanges: jest.fn(() => mockRanges),
    insertContentControl: jest.fn(() => mockContentControl),
    insertField: jest.fn(),
    insertEndnote: jest.fn(() => mockNote),
    insertPictureFromBase64: jest.fn(() => mockShape),
    insertComment: jest.fn(() => ({ content: 'Please review this project summary.' })),
    getTrackedChanges: jest.fn(() => mockTrackedChanges),
    detectLanguage: jest.fn(() => ({ value: true })),
  });
  mockRange.insertText.mockReturnValue(mockRange);
  Object.assign(mockParagraph, {
    clear: jest.fn(),
    delete: jest.fn(),
    getHtml: jest.fn(() => ({ value: '<p>Mock paragraph</p>' })),
    getOoxml: jest.fn(() => ({ value: '<w:p>Mock paragraph</w:p>' })),
    insertHtml: jest.fn(() => mockRange),
    insertOoxml: jest.fn(() => mockRange),
    insertText: jest.fn(() => mockRange),
    insertBreak: jest.fn(),
    insertInlinePictureFromBase64: jest.fn(() => mockInlinePicture),
    startNewList: jest.fn(() => mockList),
    getNext: jest.fn(() => mockParagraph),
    getPrevious: jest.fn(() => mockParagraph),
    getNextOrNullObject: jest.fn(() => ({ isNullObject: true, load: jest.fn() })),
    getPreviousOrNullObject: jest.fn(() => ({ isNullObject: true, load: jest.fn() })),
    getTextRanges: jest.fn(() => mockRanges),
  });
  // Two fixed paragraphs support location-comparison actions without modeling document mutation.
  mockParagraphs.items.push({ ...mockParagraph });
  Object.assign(mockTable, { rowCount: 2, values: [['Task', 'Status'], ['Review', 'Ready']] });
  Object.assign(mockCellBody, { getRange: jest.fn(() => mockRange) });
  Object.assign(mockBody, { font: { bold: false, color: '' }, styleBuiltIn: 'Normal', insertText: jest.fn(() => mockRange) });

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
      changeTrackingMode: 'Off',
      trackMoves: false,
      revisions: { items: [mockRevision], load: jest.fn() },
      activeWindow: mockWindow,
      windows: { items: [mockWindow], load: jest.fn() },
      addStyle: jest.fn(() => mockStyle),
      getStyles: jest.fn(() => mockStyles),
      getSelection: jest.fn(() => mockRange),
      applyQuickStyleSet: jest.fn(),
      properties: mockProperties,
      coauthoring: mockCoauthoring,
      compare: jest.fn(),
      pageSetup: mockPageSetup,
      sections: mockSections,
      tablesOfAuthorities: mockTablesOfAuthorities,
      tablesOfContents: mockTablesOfContents,
      tablesOfFigures: mockTablesOfFigures,
    },
    sync: jest.fn().mockResolvedValue(undefined),
  };

  const mockData = {
    run: jest.fn(async (contextOrCallback: Function | object, callback?: Function) => {
      if (typeof contextOrCallback !== 'function' && contextOrCallback !== mockRange && contextOrCallback !== mockContext) {
        throw new Error('Word.run received an unsupported mock object or context.');
      }
      const action = typeof contextOrCallback === 'function' ? contextOrCallback : callback;
      if (!action) {
        throw new Error('Word.run requires a callback.');
      }
      await action(mockContext);
    }),
  };

  // Create the mock object
  const wordMockObject = new OfficeMockObject(mockData, OfficeApp.Word) as any;

  // Add enums that some snippets need
  wordMockObject.CompareTarget = {
    compareTargetCurrent: 'Current',
    compareTargetNew: 'New',
  };

  wordMockObject.ChangeTrackingMode = {
    off: 'Off',
    trackAll: 'TrackAll',
    trackMineOnly: 'TrackMineOnly',
  };

  wordMockObject.RevisionType = {
    movedFrom: 'MovedFrom',
    movedTo: 'MovedTo',
  };

  wordMockObject.CoauthoringLockType = {
    reservation: 'reservation',
    ephemeral: 'ephemeral',
  };

  wordMockObject.BorderLineStyle = {
    double: 'Double',
  };

  wordMockObject.LineWidth = {
    pt150: 'Pt150',
  };

  wordMockObject.GutterPosition = {
    left: 'Left',
  };

  wordMockObject.PageOrientation = {
    portrait: 'Portrait',
  };

  wordMockObject.PaperSize = {
    letter: 'Letter',
  };

  wordMockObject.SectionStart = {
    continuous: 'Continuous',
    oddPage: 'OddPage',
  };
  wordMockObject.BevelType = {
    circle: 'circle',
  };

  wordMockObject.InsertLocation = {
    before: 'Before',
    after: 'After',
    start: 'Start',
    end: 'End',
    replace: 'Replace',
  };

  wordMockObject.BreakType = {
    page: 'Page',
    sectionContinuous: 'SectionContinuous',
    sectionEven: 'SectionEven',
    sectionNext: 'SectionNext',
    sectionOdd: 'SectionOdd',
  };

  wordMockObject.Alignment = {
    centered: 'Centered',
    justified: 'Justified',
    left: 'Left',
    right: 'Right',
  };

  wordMockObject.VerticalAlignment = {
    bottom: 'Bottom',
    center: 'Center',
    top: 'Top',
  };

  wordMockObject.RangeLocation = {
    content: 'Content',
    start: 'Start',
    end: 'End',
    whole: 'Whole',
  };

  wordMockObject.BuiltInStyleName = {
    gridTable5Dark_Accent2: 'Grid Table 5 Dark - Accent 2',
    heading1: 'Heading 1',
    heading2: 'Heading 2',
    normal: 'Normal',
    title: 'Title',
  };

  wordMockObject.ApplyQuickStyleSet = {
    sessionStart: 'SessionStart',
    template: 'Template',
  };

  wordMockObject.StyleType = {
    paragraph: 'Paragraph',
  };

  wordMockObject.BorderType = {
    dashed: 'Dashed',
    single: 'Single',
  };

  wordMockObject.BorderWidth = {
    pt100: 'Pt100',
    pt225: 'Pt225',
  };

  wordMockObject.ShadingTextureType = {
    darkTrellis: 'DarkTrellis',
    lightTrellis: 'LightTrellis',
    solid: 'Solid',
  };

  wordMockObject.TabLeader = {
    dashes: 'Dashes',
  };
  wordMockObject.SelectionMode = { end: 'End', start: 'Start', select: 'Select' };
  wordMockObject.CharacterCase = { upper: 'Upper' };
  wordMockObject.CharacterWidth = { full: 'Full' };
  wordMockObject.HorizontalInVerticalType = { fitInLine: 'FitInLine' };
  wordMockObject.Kana = { hiragana: 'Hiragana' };
  wordMockObject.TwoLinesInOneType = { parentheses: 'Parentheses' };
  wordMockObject.LanguageId = { englishUS: 'EnglishUS', japanese: 'Japanese' };
  wordMockObject.EditorType = { everyone: 'Everyone' };
  wordMockObject.FieldType = { quote: 'Quote' };
  wordMockObject.UnderlineType = { single: 'Single' };
  wordMockObject.SearchOptions = { newObject: jest.fn(() => ({ matchCase: false, matchWholeWord: false })) };

  return {
    mockObject: wordMockObject,
    mockContext,
    mockRange: extendedRange,
    mockBody,
    mockParagraph,
    mockParagraphs,
    mockParagraphFormat,
    mockShading,
    mockStyle,
    mockStyles,
    mockBreak,
    mockBreaks,
    mockPage,
    mockPages,
    mockTable,
    mockTables,
    mockTableCell,
    mockTableColumn,
    mockTableColumns,
    mockProperties,
    mockCoauthoring,
    mockCoauthor,
    mockCoauthoringLock,
    mockCoauthoringLockCollection,
    mockConflict,
    mockConflictCollection,
    mockCoauthoringUpdate,
    mockUpdateCollection,
    mockTableOfAuthorities,
    mockTableOfContents,
    mockTableOfFigures,
    mockTablesOfAuthorities,
    mockTablesOfContents,
    mockTablesOfFigures,
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
    fill: { setSolidColor: jest.fn() },
    lineFormat: { color: '' },
    textFrame: {
      textRange: {
        text: 'Mock text',
        font: { size: 9, color: '' },
      },
    },
  };

  const mockThemeColorScheme = {
    getThemeColor: jest.fn(() => ({ value: '#4472C4' })),
    setThemeColor: jest.fn(),
  };

  const mockShapes = {
    items: Array(shapeCount).fill(mockShape),
    addTextBox: jest.fn(() => mockShape),
    addShape: jest.fn(() => mockShape),
    addGeometricShape: jest.fn(() => mockShape),
    load: jest.fn(),
  };

  const mockSlide = {
    id: 'slide-1',
    shapes: mockShapes,
    themeColorScheme: mockThemeColorScheme,
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

  const pptMockObject = new OfficeMockObject(mockData, OfficeApp.PowerPoint) as any;

  pptMockObject.ThemeColor = {
    dark1: 'dark1',
    dark2: 'dark2',
    light1: 'light1',
    light2: 'light2',
    accent1: 'accent1',
    accent2: 'accent2',
    accent3: 'accent3',
    accent4: 'accent4',
    accent5: 'accent5',
    accent6: 'accent6',
    hyperlink: 'hyperlink',
    followedHyperlink: 'followedHyperlink',
  };

  pptMockObject.GeometricShapeType = {
    rectangle: 'rectangle',
    ellipse: 'ellipse',
    triangle: 'triangle',
    rightTriangle: 'rightTriangle',
    diamond: 'diamond',
    roundedRectangle: 'roundedRectangle',
  };

  return {
    mockObject: pptMockObject,
    mockContext,
    mockSlides,
    mockSlide,
    mockShapes,
    mockShape,
    mockThemeColorScheme,
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
