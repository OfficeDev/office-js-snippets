import * as path from 'path';
import * as fs from 'fs';
import * as jsyaml from 'js-yaml';
import { Dictionary } from './helpers';

import { SnippetProcessedData, banner, readDir, officeHostsToAppNames, writeFile, rmRf, mkDir } from './helpers';
import { status } from './status';
import parseXlsx from 'excel';

const SNIPPET_EXTRACTOR_METADATA_FOLDER_NAME = 'snippet-extractor-metadata';

interface MappingFileRowData {
    class: string; member: string; memberId: string, snippetId: string; snippetFunction: string
}
const headerNames: (keyof MappingFileRowData)[] =
   ['class', 'member', 'memberId', 'snippetId', 'snippetFunction'];


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
    const hostName = officeHostsToAppNames[
        filename.substr(0, filename.length - '.xlsx'.length).toUpperCase()];

    banner(`Extracting reference-doc snippet bits for ${hostName}`);

    const lines: MappingFileRowData[] =
        await new Promise((resolve: (data: MappingFileRowData[]) => void, reject) => {
            const fullFilePath = path.join(
                path.resolve(SNIPPET_EXTRACTOR_METADATA_FOLDER_NAME),
                filename
            );
            parseXlsx(fullFilePath).then((data) => {
                if (data.length < 2) {
                    reject(new Error('No data rows found'));
                }

                if (data[0].length !== headerNames.length) {
                    reject(
                        new Error('Unexpected number of columns. Expecting the following ' +
                            headerNames.length + ' columns: ' +
                            headerNames.map(name => `"${name}"`).join(', ')
                        )
                    );
                }

                // Remove the first line, since it's the header line
                data.splice(0, 1);

                resolve(data.map((row: string[]) => {
                    if (row.find(text => text.startsWith('//'))) {
                        return null;
                    }

                    let result: MappingFileRowData = {} as any;
                    row.forEach((column: string, index) => {
                        result[headerNames[index]] = column;
                    });
                    return result;
                }).filter(item => item));
            });
        });

    const allSnippetData: { [key: string]: string[] } = {};

    lines.map(row => getExtractedDataFromSnippet(row, snippetIdsToFilenames, accumulatedErrors))
        .filter(item => item)
        .forEach((text, index) => {
            const row = lines[index];
            let fullName = `${hostName}.${row.class.trim()}#${row.member.trim()}:member`;
            if (row.memberId) {
                fullName += `(${row.memberId})`;
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
            const functionDeclarationLine = fullSnippetTextArray[arrayIndex];
            const functionHasNoParams = functionDeclarationLine.indexOf(targetText + ')') >= 0;

            const spaceFollowedByWordsRegex = /^(\s*)(.*)$/;
            const preWhitespaceCount = spaceFollowedByWordsRegex.exec(functionDeclarationLine)[1].length;
            const targetClosingText = ' '.repeat(preWhitespaceCount) + '}';
            fullSnippetTextArray.splice(0, arrayIndex + (functionHasNoParams ? 1 : 0));

            const closingIndex = fullSnippetTextArray.findIndex(text => text.indexOf(targetClosingText) === 0);
            if (closingIndex < 0) {
                throw new Error(`Could not find a closing bracket at same level of indent as the original function declaration ("${targetText}")`);
            }

            const indented = fullSnippetTextArray.slice(0, closingIndex + (functionHasNoParams ? 0 : 1));
            const whitespaceCountOnFirstLine = spaceFollowedByWordsRegex.exec(fullSnippetTextArray[0])[1].length;

            // Place snippet location as comment.
            const editedFilename = filename.substr(filename.lastIndexOf('samples')).replace(/\\/g, '/');
            text = '// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/master/' + editedFilename + '\n';

            text += indented
                .map(line => {
                    if (line.substr(0, whitespaceCountOnFirstLine).trim().length === 0) {
                        return line.substr(whitespaceCountOnFirstLine);
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
