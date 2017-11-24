import * as path from 'path';
import * as fs from 'fs';
import * as csv from 'csv-parser';

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
    if (!filename.endsWith('.csv')) {
        throw new Error(`Expecting ${filename} to end in ".csv"`);
    }

    const hostName = officeHostsToAppNames[
        filename.substr(0, filename.length - '.csv'.length).toUpperCase()];

    banner(`Extracting reference-doc snippet bits for ${hostName}`);

    const lines: MappingFileRowData[] =
        await new Promise((resolve: (data: MappingFileRowData[]) => void, reject) => {
            const parser = csv(headerNames);
            const fullFilePath = path.join(
                path.resolve(SNIPPET_EXTRACTOR_METADATA_FOLDER_NAME),
                filename
            );

            const lines: MappingFileRowData[] = [];
            fs.createReadStream(fullFilePath)
                .pipe(parser)
                .on('data', (data: MappingFileRowData) => lines.push(data))
                .on('error', reject)
                .on('end', () => {
                    // Remove the first line, since it's the header line
                    lines.splice(0, 1);

                    resolve(lines);
                });
        });

    const allSnippetData: { [key: string]: string[] } = {};

    (await Promise.all(lines.map(row => getExtractedDataFromSnippet(row))))
        .forEach(data => {
            const { text } = data;
            const fullName = hostName + '.' + data.name;
            if (!allSnippetData[fullName]) {
                allSnippetData[fullName] = [];
            }
            allSnippetData[fullName].push(text);
        });

    return allSnippetData;
}

async function getExtractedDataFromSnippet(
    row: MappingFileRowData
): Promise<{ name: string, text: string }> {
    return Promise.resolve(null);
}
