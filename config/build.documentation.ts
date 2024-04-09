import * as path from 'path';
import * as fs from 'fs';
import * as jsyaml from 'js-yaml';
import { Dictionary } from './helpers';

import { SnippetProcessedData, readDir, writeFile, rmRf, mkDir } from './helpers';
import { status } from './status';
const ExcelJS = require('exceljs');

const SNIPPET_EXTRACTOR_METADATA_FOLDER_NAME = 'snippet-extractor-metadata';

interface MappingFileRowData {
    package: string, class: string; member: string; memberId: string, snippetId: string; snippetFunction: string
}
const headerNames: (keyof MappingFileRowData)[] =
   ['package', 'class', 'member', 'memberId', 'snippetId', 'snippetFunction'];


export async function buildReferenceDocSnippetExtracts(
    snippets: Dictionary<SnippetProcessedData>,
    accumulatedErrors: Array<string | Error>
): Promise<void> {
    let files = (await readDir(path.resolve(SNIPPET_EXTRACTOR_METADATA_FOLDER_NAME)))
        .filter(name => name.endsWith('.xlsx'))
        .filter(name => !name.startsWith('~$'));

    const snippetIdsToFilenames: { [key: string]: string } = {};
    snippets.values().forEach(item => {
        snippetIdsToFilenames[item.id] = item.fullPath;
    });

    let snippetExtractsPerHost = await Promise.all(
        files.map(file => buildSnippetExtractsPerHost(file, snippetIdsToFilenames, accumulatedErrors))
    );

    await rmRf('snippet-extractor-output');
    await mkDir('snippet-extractor-output');

    const contents = snippetExtractsPerHost.map(extracts => jsyaml.safeDump(extracts)).join('');
    await writeFile(path.resolve(`snippet-extractor-output/snippets.yaml`), contents);
}

async function buildSnippetExtractsPerHost(
    filename: string,
    snippetIdsToFilenames: { [key: string]: string },
    accumulatedErrors: Array<string | Error>
): Promise<{ [key: string]: string[] }> {

    const lines: MappingFileRowData[] =
        await new Promise(async (resolve: (data: MappingFileRowData[]) => void, reject) => {
            const fullFilePath = path.join(
                path.resolve(SNIPPET_EXTRACTOR_METADATA_FOLDER_NAME),
                filename
            );
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(fullFilePath);
            const worksheet = workbook.worksheets[0];
            if (worksheet.rowCount < 2) {
                reject(new Error('No data rows found'));
            }

            if (worksheet.getRow(1).cellCount !== headerNames.length) {
                reject(
                    new Error('Unexpected number of columns. Expecting the following ' +
                        headerNames.length + ' columns: ' +
                        headerNames.map(name => `"${name}"`).join(', ')
                    )
                );
            }

            let mappedRowData: MappingFileRowData[] = [];
            worksheet.eachRow((row, rowNumber) => {
                if (rowNumber !== 1 && !row.getCell(1).value.startsWith('//')) {
                    let result: MappingFileRowData = {} as any;
                    row.eachCell((cell, index) => {
                        result[headerNames[index - 1]] = cell.value;
                    });
                    mappedRowData.push(result);
                }
            });
            resolve(mappedRowData.filter(item => item));
        });

    const allSnippetData: { [key: string]: string[] } = {};

    lines.map(row => getExtractedDataFromSnippet(row, snippetIdsToFilenames, accumulatedErrors))
        .filter(item => item)
        .forEach((text, index) => {
            const row = lines[index];
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
): string {
    const updatingStatusText = `${row.class}.${row.member}: function "${row.snippetFunction}" from snippet ID "${row.snippetId}"`;
    status.add(updatingStatusText);
    let text: string;

    const filename = snippetIdsToFilenames[row.snippetId];
    if (filename) {
        try {
            const script = (jsyaml.safeLoad(fs.readFileSync(filename).toString()) as ISnippet).script.content;

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
            const preWhitespaceCount = spaceFollowedByWordsRegex.exec(functionDeclarationLine)[1].length;
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
            const whitespaceCountOnFirstLine = spaceFollowedByWordsRegex.exec(fullSnippetTextArray[0])[1].length;

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
            accumulatedErrors.push(`${row.snippetId}: ${exception.message || exception}`);
        }
    } else {
        accumulatedErrors.push(`Could not find snippet id "${row.snippetId}" in mapping table`);
    }

    status.complete(text ? true : false /*succeeded */, updatingStatusText);
    return text;
}
