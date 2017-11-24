import * as path from 'path';

import { Dictionary } from '@microsoft/office-js-helpers';
import { SnippetProcessedData, banner, readDir, officeHostsToAppNames } from './helpers';

const SNIPPET_EXTRACTOR_METADATA_FOLDER_NAME = 'snippet-extractor-metadata';

interface MappingFileRowData {
    class: string; member: string; snippetId: string; functionName: string
}
const headerNames: (keyof MappingFileRowData)[] =
    ['class', 'member', 'snippetId', 'functionName'];


export async function buildReferenceDocSnippetExtracts(snippets: Dictionary<SnippetProcessedData>) {
    let files = await readDir(path.resolve(SNIPPET_EXTRACTOR_METADATA_FOLDER_NAME));
    return Promise.all(files.map(file => buildSnippetExtractsPerHost(file)));
}

async function buildSnippetExtractsPerHost(filename: string) {
    if (!filename.endsWith('.xlsx')) {
        throw new Error(`Expecting ${filename} to end in ".xlsx"`);
    }

    const hostName = officeHostsToAppNames[
        filename.substr(0, filename.length - '.xlsx'.length).toUpperCase()];

    banner(`Extracting reference-doc snippet bits for ${hostName}`);

    const lines: MappingFileRowData[] =
        await new Promise((resolve: (data: MappingFileRowData[]) => void, reject) => {
            const parseXlsx = require('excel');

            const fullFilePath = path.join(
                path.resolve(SNIPPET_EXTRACTOR_METADATA_FOLDER_NAME),
                filename
            );
            parseXlsx(fullFilePath, (err, rows: any[][]) => {
                if (err) {
                    reject(err);
                }

                if (rows.length < 2) {
                    reject(new Error('No data rows found'));
                }

                if (rows[0].length !== headerNames.length) {
                    reject(
                        new Error('Unexpected number of columns. Expecting the following ' +
                            headerNames.length + ' columns: ' +
                            headerNames.map(name => `"${name}"`).join(', ')
                        )
                    );
                }

                // Remove the first line, since it's the header line
                rows.splice(0, 1);

                resolve(
                    rows.map(row => {
                        let result: MappingFileRowData = {} as any;
                        row.forEach((column, index) => {
                            result[headerNames[index]] = column;
                        });
                        return result;
                    })
                );

                resolve(lines);
            });
        });

    lines.forEach(item => console.log(JSON.stringify(item)));
    // const allSnippetData: { [key: string]: string[] } = {};

    // (await Promise.all(lines.map(row => getExtractedDataFromSnippet(row))))
    //     .forEach(data => {
    //         const { text } = data;
    //         const fullName = hostName + '.' + data.name;
    //         if (!allSnippetData[fullName]) {
    //             allSnippetData[fullName] = [];
    //         }
    //         allSnippetData[fullName].push(text);
    //     });

    // return allSnippetData;
}

// async function getExtractedDataFromSnippet(
//     row: MappingFileRowData
// ): Promise<{ name: string, text: string }> {
//     return Promise.resolve(null);
// }
