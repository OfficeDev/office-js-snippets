import * as path from 'path';
import * as fs from 'fs';
import * as jsyaml from 'js-yaml';
import { Dictionary } from './helpers';

import { SnippetProcessedData, readDir, writeFile, rmRf, mkDir } from './helpers';
import { status } from './status';

const SNIPPET_EXTRACTOR_METADATA_FOLDER_NAME = 'snippet-extractor-metadata';

interface MappingFileRowData {
    package: string, class: string; member: string; memberId: string, snippetId: string; snippetFunction: string
}

/** Maps CSV column headers to MappingFileRowData property names. */
const csvHeaderToFieldName: { [csvHeader: string]: keyof MappingFileRowData } = {
    'Package': 'package',
    'Class': 'class',
    'Member Name': 'member',
    'Member ID or top-level category': 'memberId',
    'SnippetIdInTheYAMLFile': 'snippetId',
    'MethodNameInTheSnippet': 'snippetFunction',
};

function parseCsvLine(line: string): string[] {
    const fields: string[] = [];
    let current = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
            if (inQuotes && line[i + 1] === '"') {
                current += '"';
                i++;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (ch === ',' && !inQuotes) {
            fields.push(current);
            current = '';
        } else {
            current += ch;
        }
    }
    fields.push(current);
    return fields;
}

export async function buildReferenceDocSnippetExtracts(
    snippets: Dictionary<SnippetProcessedData>,
    accumulatedErrors: Array<string | Error>
): Promise<void> {
    let files = (await readDir(path.resolve(SNIPPET_EXTRACTOR_METADATA_FOLDER_NAME)))
        .filter(name => name.endsWith('.csv'));

    const snippetIdsToFilenames: { [key: string]: string } = {};
    snippets.values().forEach(item => {
        snippetIdsToFilenames[item.id] = item.fullPath;
    });

    let snippetExtractsPerHost = await Promise.all(
        files.map(file => buildSnippetExtractsPerHost(file, snippetIdsToFilenames, accumulatedErrors))
    );

    await rmRf('snippet-extractor-output');
    await mkDir('snippet-extractor-output');

    const contents = snippetExtractsPerHost.map(extracts => jsyaml.dump(extracts)).join('');
    await writeFile(path.resolve(`snippet-extractor-output/snippets.yaml`), contents);
}

async function buildSnippetExtractsPerHost(
    filename: string,
    snippetIdsToFilenames: { [key: string]: string },
    accumulatedErrors: Array<string | Error>
): Promise<{ [key: string]: string[] }> {

    const fullFilePath = path.join(
        path.resolve(SNIPPET_EXTRACTOR_METADATA_FOLDER_NAME),
        filename
    );
    const rawLines = fs.readFileSync(fullFilePath, 'utf8')
        .replace(/^\uFEFF/, '') // strip UTF-8 BOM if present
        .split('\n')
        .map(line => line.replace(/\r$/, ''))
        .filter(line => line.trim().length > 0);

    if (rawLines.length < 2) {
        throw new Error(`No data rows found in ${filename}`);
    }

    const csvHeaders = parseCsvLine(rawLines[0]);
    const columnIndices: { [field in keyof MappingFileRowData]?: number } = {};
    csvHeaders.forEach((header, index) => {
        const fieldName = csvHeaderToFieldName[header];
        if (fieldName) {
            columnIndices[fieldName] = index;
        }
    });

    const expectedFields = Object.keys(csvHeaderToFieldName) as (keyof typeof csvHeaderToFieldName)[];
    const missingFields = expectedFields.filter(h => csvHeaderToFieldName[h] && columnIndices[csvHeaderToFieldName[h]] === undefined);
    if (missingFields.length > 0) {
        throw new Error(`Missing expected columns in ${filename}: ${missingFields.join(', ')}`);
    }

    const lines: MappingFileRowData[] = rawLines.slice(1)
        .filter(line => !parseCsvLine(line)[0].startsWith('//'))
        .map(line => {
            const cells = parseCsvLine(line);
            const result: MappingFileRowData = {} as any;
            (Object.keys(columnIndices) as (keyof MappingFileRowData)[]).forEach(field => {
                result[field] = cells[columnIndices[field]!] ?? '';
            });
            return result;
        });

    const allSnippetData: { [key: string]: string[] } = {};

    lines.forEach(row => {
        const text = getExtractedDataFromSnippet(row, snippetIdsToFilenames, accumulatedErrors);
        if (!text) { return; }

        let hostName = row.package;
        let fullName;
        if (row.member) { /* If the mapping is for a field */
            fullName = `${hostName}.${row.class.trim()}#${row.member.trim()}:member`;
            if (row.memberId) {
                fullName += `(${row.memberId})`;
            }
        } else { /* If the mapping is for a top-level sample (like an enum) */
            fullName = `${hostName}.${row.class.trim()}:${row.memberId.trim()}`;
        }

        if (!allSnippetData[fullName]) {
            allSnippetData[fullName] = [];
        }
        allSnippetData[fullName].push(text);
    });
    return allSnippetData;
}

function getExtractedDataFromSnippet(
    row: MappingFileRowData,
    snippetIdsToFilenames: { [key: string]: string },
    accumulatedErrors: Array<string | Error>
): string | undefined {
    const updatingStatusText = `${row.class}.${row.member}: function "${row.snippetFunction}" from snippet ID "${row.snippetId}"`;
    status.add(updatingStatusText);
    let text: string | undefined;

    const filename = snippetIdsToFilenames[row.snippetId];
    if (filename) {
        try {
            const script = (jsyaml.load(fs.readFileSync(filename).toString()) as ISnippet).script?.content ?? '';

            const fullSnippetTextArray = script.split('\n')
                .map(line => line.replace(/\r/, ''));
            const targetText = `function ${row.snippetFunction}(`;

            let arrayIndex = fullSnippetTextArray.findIndex(text => text.indexOf(targetText) >= 0);
            if (arrayIndex < 0) {
                throw new Error(`Invalid entry in the metadata mapping file -- snippet function "${row.snippetFunction}" does not exist within snippet "${filename}"`);
            }

            let jsDocCommentIndex = -1;
            if (arrayIndex > 0 && fullSnippetTextArray[arrayIndex - 1].indexOf('*/') >= 0) {
                for (let i = arrayIndex - 1; i >= 0; i--) {
                    if (fullSnippetTextArray[i].indexOf('/**') >= 0) {
                        jsDocCommentIndex = i;
                        break;
                    }
                }
            }

            const functionDeclarationLine = fullSnippetTextArray[arrayIndex];
            const functionHasNoParams = functionDeclarationLine.indexOf(targetText + ')') >= 0;

            const spaceFollowedByWordsRegex = /^(\s*)(.*)$/;
            const preWhitespaceCount = spaceFollowedByWordsRegex.exec(functionDeclarationLine)![1].length;
            const targetClosingText = ' '.repeat(preWhitespaceCount) + '}';
            if (jsDocCommentIndex >= 0) {
                fullSnippetTextArray.splice(0, jsDocCommentIndex);
            } else {
                fullSnippetTextArray.splice(0, arrayIndex + (functionHasNoParams ? 1 : 0));
            }

            const closingIndex = fullSnippetTextArray.findIndex(text => text.indexOf(targetClosingText) === 0);
            if (closingIndex < 0) {
                throw new Error(`Could not find a closing bracket at same level of indent as the original function declaration ("${targetText}")`);
            }

            const indented = fullSnippetTextArray.slice(0, closingIndex + (functionHasNoParams ? 0 : 1));
            const whitespaceCountOnFirstLine = spaceFollowedByWordsRegex.exec(fullSnippetTextArray[0])![1].length;

            // Place snippet location as comment.
            const editedFilename = filename.substring(filename.lastIndexOf('samples')).replace(/\\/g, '/');
            text = '// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/' + editedFilename + '\n\n';

            text += indented
                .map(line => {
                    if (line.substring(0, whitespaceCountOnFirstLine).trim().length === 0) {
                        return line.substring(whitespaceCountOnFirstLine);
                    } else {
                        return line;
                    }
                })
                .join('\n');
        }
        catch (exception) {
            accumulatedErrors.push(`${row.snippetId}: ${(exception as any).message || exception}`);
        }
    } else {
        accumulatedErrors.push(`Could not find snippet id "${row.snippetId}" in mapping table`);
    }

    status.complete(text ? true : false /*succeeded */, updatingStatusText);
    return text;
}
