#!/usr/bin/env node --harmony

import * as path from 'path';
import { isNil, isString, isArray, isEmpty, sortBy, cloneDeep } from 'lodash';
import * as chalk from 'chalk';
import { status } from './status';
import {
    SnippetFileInput, SnippetProcessedData,
    followsNamingGuidelines, isCUID,
    rmRf, mkDir, readDir, getFiles, writeFile, banner, getPrintableDetails, Dictionary
} from './helpers';
import { buildReferenceDocSnippetExtracts } from './build.documentation';
import { getShareableYaml } from './snippet.helpers';
import { processLibraries } from './libraries.processor';
import { startCase, groupBy, map } from 'lodash';
import * as jsyaml from 'js-yaml';
import escapeStringRegexp = require('escape-string-regexp');
import * as fsx from 'fs-extra';


const PRIVATE_SAMPLES = 'private-samples';
const PUBLIC_SAMPLES = 'samples';
const snippetFilesToUpdate: Array<{ path: string; contents: string }> = [];
const accumulatedErrors: Array<string | Error> = [];
const sortingCriteria = ['group', 'order', 'id'];

const officeHosts = ['ACCESS', 'EXCEL', 'ONENOTE', 'OUTLOOK', 'POWERPOINT', 'PROJECT', 'WORD'];
const defaultApiSets = {
    'EXCEL': {
        'ExcelApi': 1.1
    },
    'ONENOTE': {
        'OneNoteApi': 1.1
    },
    'WORD': {
        'WordApi': 1.1
    }

    // any other host is allowed to have no API sets specified
};


(async () => {
    let processedSnippets = new Dictionary<SnippetProcessedData>();
    await Promise.resolve()
        .then(() => processSnippets(processedSnippets))
        .then(updateModifiedFiles)
        .then(() => checkSnippetsForUniqueIDs(processedSnippets))
        .then(() => generatePlaylists(processedSnippets))
        .then(copyAndUpdatePlaylistFolders)
        .then(() => buildReferenceDocSnippetExtracts(processedSnippets, accumulatedErrors))
        .then(() => {
            if (accumulatedErrors.length > 0) {
                throw accumulatedErrors;
            }
        })
        .then(() => {
            banner('Done!', null, chalk.bold.green);
        })
        .catch(handleError);

    process.exit(0);
})();


async function processSnippets(processedSnippets: Dictionary<SnippetProcessedData>) {
    banner('Loading & processing snippets');
    let files: SnippetFileInput[] = []
        .concat(getFiles(path.resolve(PRIVATE_SAMPLES)))
        .concat(getFiles(path.resolve(PUBLIC_SAMPLES)));

    (await Promise.all(files.map(file => processAndValidateSnippet(file))))
        .filter(file => file !== null)
        .map(file => processedSnippets.set(file.rawUrl, file));


    // Helpers:

    async function processAndValidateSnippet(file: SnippetFileInput): Promise<SnippetProcessedData> {
        const messages: Array<string | Error> = [];
        try {
            status.add(`Processing ${file.relativePath}`);
            let dir = file.isPublic ? PUBLIC_SAMPLES : PRIVATE_SAMPLES;

            const fullPath = path.resolve(dir, file.relativePath);
            const originalFileContents = fsx.readFileSync(fullPath).toString().trim();
            let snippet = jsyaml.safeLoad(originalFileContents) as ISnippet;

            // Do validations & auto-corrections
            validateStringFieldNotEmptyOrThrow(snippet, 'name');
            validateStringFieldNotEmptyOrThrow(snippet, 'description');
            validateId(snippet, file.relativePath, messages);
            validateSnippetHost(snippet, file.host, messages);
            validateAtTypesDeclarations(snippet, messages);
            validateOfficialOfficeJs(snippet, file.host, file.group, messages);
            validateApiSetNonEmpty(snippet, file.host, file.relativePath, messages);
            validateVersionNumbersOnLibraries(snippet, messages);
            validateTabsInsteadOfSpaces(snippet, messages);
            validateProperFabric(snippet);

            // Additional fields relative to what is normally exposed in sharing
            // (and/or that would normally get erased when doing an export):
            const additionalFields: ISnippet = <any>{};
            additionalFields.id = snippet.id;
            additionalFields.api_set = snippet.api_set;
            additionalFields.author = snippet.author;

            if ((typeof (snippet as any).order) !== 'undefined') {
                // # for ordering, if present (used for samples only)
                (additionalFields as any).order = (snippet as any).order;
            }

            // Finally, some fields simply don't apply, and should be deleted.
            delete snippet.gist;

            let finalFileContents = getShareableYaml(snippet, additionalFields).trim();

            let isDifferent = finalFileContents.replace(/\r\n/g, '\n') !== originalFileContents.replace(/\r\n/g, '\n');
            if (isDifferent) {
                messages.push(chalk.bold.yellow('Final snippet != original snippet. Queueing to write in new changes.'));
                snippetFilesToUpdate.push({ path: fullPath, contents: finalFileContents });
            }
            status.complete(true /*success*/, `Processing ${file.relativePath}`, messages);

            const rawUrl = `https://raw.githubusercontent.com/` +
                `OfficeDev/office-js-snippets/main` +
                `/${dir}/${file.host}/${file.group}/${file.file_name}`;

            if (messages.findIndex(item => item instanceof Error) >= 0) {
                accumulatedErrors.push(`One or more critical errors on ${file.relativePath}`);
            }

            // Define dictionary of words in file.group that require special casing
            let dictionary = {
                'Apis': 'APIs',
                'Pivottable': 'PivotTable',
                'Xml': 'XML'
            };

            let groupName = replaceUsingDictionary(dictionary, startCase(file.group));

            return {
                id: snippet.id,
                name: snippet.name,
                fileName: file.file_name,
                relativePath: file.relativePath,
                fullPath: file.fullPath,
                description: snippet.description,
                host: file.host,
                rawUrl: rawUrl,
                group: groupName,
                order: (typeof (snippet as any).order === 'undefined') ? 100 /* nominally 100 */ : (snippet as any).order,
                api_set: snippet.api_set,
                isPublic: file.isPublic
            };
        } catch (exception) {
            messages.push(exception);
            status.complete(false /*success*/, `Processing ${file.relativePath}`, messages);
            accumulatedErrors.push(`Failed to process ${file.relativePath}: ${exception.message || exception}`);
            return null;
        }


        function replaceUsingDictionary(dictionary: { [key: string]: string }, originalName: string): string {
            let text = startCase(file.group);
            let parts = text.split(' ').map(item => dictionary[item] || item);
            return parts.join(' ');
        }
    }

    function validateProperFabric(snippet: ISnippet): void {
        const libs = snippet.libraries.split('\n').map(reference => reference.trim());
        if (libs.indexOf('office-ui-fabric-core@11.1.0/dist/css/fabric.min.css') >= 0) {
            if (libs.indexOf('office-ui-fabric-js@1.5.0/dist/css/fabric.components.min.css') <= 0) {
                throw new Error('Fabric reference is specified, without a reference to a corresponding "fabric.components.min.css". Please add this second Fabric reference as well.');
            }
        }
    }

    function validateStringFieldNotEmptyOrThrow(snippet: ISnippet, field: string): void {
        if (isNil(snippet[field])) {
            throw `Snippet ${field} may not be empty`;
        }

        if (!isString(snippet[field])) {
            throw `Snippet ${field} must be a string`;
        }

        snippet[field] = snippet[field].trim();
    }

    function validateSnippetHost(snippet: ISnippet, host: string, messages: any[]): void {
        host = host.toUpperCase();

        if (typeof snippet.host === 'undefined') {
            messages.push(`Snippet is missing "host" property. Settings based on file path to ${host}`);
            snippet.host = host;
        }

        if (snippet.host !== snippet.host.toUpperCase()) {
            messages.push(`Snippet host is inconsistently-cased. Changing to all-caps`);
            snippet.host = snippet.host.toUpperCase();
        }

        if (snippet.host !== host) {
            throw new Error(`Snippet's specified host "${snippet.host}" is different than the directory path host "${host}". Please fix the mismatch.`);
        }
    }

    function validateAtTypesDeclarations(snippet: ISnippet, messages: any[]) {
        snippet.libraries.split('\n')
            .map(reference => reference.trim())
            .filter(reference => reference.match(/^dt~.*$/gi))
            .map(reference => {
                const atTypesNotation = `@types/${reference.substr('dt~'.length)}`;
                snippet.libraries = snippet.libraries.replace(reference, atTypesNotation);
                messages.push(`Replacing reference "${reference}" with the @types notation: "${atTypesNotation}"`);
            });
    }

    function validateOfficialOfficeJs(snippet: ISnippet, host: string, group: string, messages: any[]): void {
        const isOfficeSnippet = officeHosts.indexOf(host.toUpperCase()) >= 0;
        const canonicalOfficeJsReference = 'https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js';
        const betaOfficeJsReference = 'https://appsforoffice.microsoft.com/lib/beta/hosted/office.js';
        const officeDTS = '@types/office-js';
        const betaOfficeDTS = '@types/office-js-preview';

        const officeJsReferences =
            snippet.libraries.split('\n')
                .map(reference => reference.trim())
                .filter(reference => reference.match(/^http.*\/office\.js$/gi));

        const officeDtsReferences =
            snippet.libraries.split('\n')
                .map(reference => reference.trim())
                .filter(reference => reference.match(/.*((@types\/office-js(-preview)?)|(office\.d\.ts))$/gi));
        /* Note: regex matches:
            - @types/office-js
            - @types/office-js-preview
            - https://unpkg.com/etc/office.d.ts
           But not:
            - @types/office-jsfake
            - https://unpkg.com/etc/office.d.ts.ish
            - office.d.ts.unrelated
         */

        if (!isOfficeSnippet) {
            if (officeJsReferences.length > 0 || officeDtsReferences.length > 0) {
                throw new Error(`Snippet for host "${host}" should not have a reference to either office.js or to office.d.ts`);
            }
            return;
        }


        // From here on out, can assume that is an Office snippet

        if (officeJsReferences.length === 0) {
            throw new Error(`Snippet for host "${host}" should have a reference to office.js`);
        }
        if (officeDtsReferences.length === 0) {
            throw new Error(`Snippet for host "${host}" should have a reference to office.d.ts`);
        }

        if (officeJsReferences.length > 1 || officeDtsReferences.length > 1) {
            throw new Error(`Cannot have more than one reference to office.js or to office.d.ts`);
        }

        let snippetOfficeReferenceIsOk =
            officeJsReferences[0] === canonicalOfficeJsReference ||
            (group.indexOf('preview-apis') >= 0 && officeJsReferences[0] === betaOfficeJsReference);

        if (!snippetOfficeReferenceIsOk) {
            throw new Error(`Office.js reference "${officeJsReferences[0]}" does match the canonical form of "${canonicalOfficeJsReference}" and does match any of the exceptions defined by "snippetOfficeReferenceIsOk".`);
        }


        let officeJsDtsForSameLocation = officeJsReferences[0].substr(0,
            officeJsReferences[0].length - 'office.js'.length) + 'office.d.ts';

        if (officeJsReferences[0] === canonicalOfficeJsReference) {
            let isCorrectCorrespondingDts = officeDtsReferences[0] === officeDTS ||
                officeDtsReferences[0] === officeJsDtsForSameLocation;
            if (!isCorrectCorrespondingDts) {
                throw new Error(`Office.js reference is "${officeJsReferences[0]}" so the types reference should be "${officeDTS}" or the "office.d.ts" from the same location as "office.js".`);
            }
        }

        if (officeJsReferences[0] === betaOfficeJsReference) {
            let isCorrectCorrespondingDts = officeDtsReferences[0] === betaOfficeDTS ||
                officeDtsReferences[0] === officeJsDtsForSameLocation;
            if (!isCorrectCorrespondingDts) {
                throw new Error(`Office.js reference is "${officeJsReferences[0]}" so the types reference should be "${betaOfficeDTS}" or the "office.d.ts" from the same location as "office.js".`);
            }
        }
    }

    function validateApiSetNonEmpty(snippet: ISnippet, host: string, localPath: string, messages: any[]): void {
        host = host.toUpperCase();

        if (typeof snippet.api_set === 'undefined') {
            snippet.api_set = {};
        }

        if (isEmpty(snippet.api_set)) {
            if (typeof defaultApiSets[host] === 'undefined') {
                // No API set required (not a host with host-specific APIs), so just exit the function
                return;
            }

            messages.push(new Error(`No API set specified. If building locally, substituting with a default of ` +
                `"${JSON.stringify(defaultApiSets[host])}", but failing the build.`));
            messages.push(new Error('   Please check your pending changes to see the substituted version.'));

            snippet.api_set = cloneDeep(defaultApiSets[host]);
        }
    }

    function validateVersionNumbersOnLibraries(snippet: ISnippet, messages: any[]): void {
        const { scriptReferences, linkReferences } = processLibraries(snippet);
        const allWithoutVersionNumbers = ([].concat(scriptReferences).concat(linkReferences) as string[])
            .filter(item => item.startsWith('https://unpkg.com/'))
            .map(item => item.substr('https://unpkg.com/'.length))
            .filter(item => {
                const containsVersionNumberRegex = /^(@[a-zA-Z_\-0-9]+\/)?([a-zA-Z_0-9\-]+)@[0-9\.]*.*$/;
                /* Tested with:
                        @microsoft/office-js-helpers
                            => wrong

                        @microsoft/office-js-helpers/lib.js
                            => wrong

                        @microsoft/office-js-helpers@0.6.5
                            => right

                        @microsoft/office-js-helpers@0.6.5
                            => right

                        jquery@0.6.0/lib.js
                            => right

                        jquery@0.6.0
                            => right

                        jquery/lib.js
                            => wrong

                        foo-bar/
                            => wrong

                        foo-bar@1.5
                            => right

                        foobar2@1.5
                            => right
                */

                return !item.match(containsVersionNumberRegex);
            });

        if (allWithoutVersionNumbers.length === 0) {
            return;
        }

        const defaultSubstitutions = {
            'jquery': 'jquery@3.1.1',
            'office-ui-fabric-js/dist/js/fabric.min.js': 'office-ui-fabric-js@1.5.0/dist/js/fabric.min.js',
            '@microsoft/office-js-helpers/dist/office.helpers.min.js': '@microsoft/office-js-helpers@0.7.4/dist/office.helpers.min.js',
            'core-js/client/core.min.js': 'core-js@2.4.1/client/core.min.js',
            'office-ui-fabric-core/dist/css/fabric.min.css': 'office-ui-fabric-core@11.1.0/dist/css/fabric.min.css',
            'office-ui-fabric-js/dist/css/fabric.min.css': 'office-ui-fabric-core@11.1.0/dist/css/fabric.min.css',
            'office-ui-fabric-js/dist/css/fabric.components.min.css': 'office-ui-fabric-js@1.5.0/dist/css/fabric.components.min.css'
        };

        let hadDefaultSubstitution = false;
        allWithoutVersionNumbers.forEach(item => {
            if (defaultSubstitutions[item]) {
                const regex = new RegExp(`(^\s*)${escapeStringRegexp(item)}(.*)`, 'm');
                snippet.libraries = snippet.libraries.replace(regex, ($0, $1, $2) => $1 + defaultSubstitutions[item] + $2);

                messages.push(new Error(`Missing version number on library ${item}. If building locally, substituting with a default of ` +
                    `"${defaultSubstitutions[item]}", but failing the build.`));
                hadDefaultSubstitution = true;
            }

            messages.push(new Error(`Missing version number on library ${item}. A version # is required for NPM packages.`));
        });

        if (hadDefaultSubstitution) {
            messages.push(new Error('Please check your pending changes to see the default-substituted library version(s).'));
        }
    }

    function validateId(snippet: ISnippet, localPath: string, messages: any[]): void {
        let originalId = snippet.id;

        // Don't want empty IDs -- or GUID-y IDs either, since they're not particularly memorable...
        if (isNil(snippet.id) || snippet.id.trim().length === 0 || isCUID(snippet.id)) {
            snippet.id = localPath;
        }

        snippet.id = snippet.id.trim().toLowerCase()
            .replace(/[^0-9a-zA-Z]/g, '-') /* replace any non-alphanumeric with a hyphen */
            .replace(/-+/g, '-') /* and ensure that don't end up with -- or --, just a single hyphen */
            .replace(/yaml$/i, '') /* remove "yaml" suffix (the ".", now "-", will get removed via hyphen-trimming below) */
            .replace(/^-+/, '') /* trim any hyphens before */
            .replace(/-+$/, '') /* and trim any at the end, as well */
            .replace(/-(\d+-)(.*)/, '-$2') /* remove any numeric prefixes like "word\01-basics\foo", replacing with "word\basics\foo" */
            .replace('-preview-apis-', '-');

        if (snippet.id !== originalId) {
            messages.push(`Snippet ID needs correcting. Replacing with an ID based on name: "${snippet.id}"`);
        }

        if (!followsNamingGuidelines(snippet.id)) {
            messages.push(new Error('Snippet ID does not follow naming conventions (only lowercase letters, numbers, and hyphens). Please change it.'));
        }
    }

    function validateTabsInsteadOfSpaces(snippet: ISnippet, messages: any[]): void {
        let replacedTabs = false;
        ['template', 'script', 'style'].forEach(fieldName => {
            if (snippet[fieldName]) {
                if (snippet[fieldName].content.indexOf('\t') >= 0) {
                    replacedTabs = true;
                    snippet[fieldName].content = snippet[fieldName].content.replace(/\t/g, '    ');
                }
            }
        });

        if (snippet.libraries.indexOf('\t') >= 0) {
            replacedTabs = true;
            snippet.libraries = snippet.libraries.replace(/\t/g, '    ');
        }

        if (replacedTabs) {
            messages.push('Snippet had one or more fields (template/script/style/libraries) ' +
                'that contained tabs instead of spaces. Replacing everything with 4 spaces.');
        }
    }
}

async function updateModifiedFiles() {
    const hasChanges = snippetFilesToUpdate.length > 0;
    banner('Updating modified files',
        hasChanges ? null : '<No files to modify>',
        hasChanges ? chalk.bold.yellow : null);

    const fileWriteRequests = [];
    snippetFilesToUpdate.forEach(item => {
        fileWriteRequests.push(
            Promise.resolve()
                .then(async () => {
                    const updatingStatusText = `Updating ${item.path}`;
                    status.add(updatingStatusText);
                    await writeFile(item.path, item.contents);
                    status.complete(true /*succeeded*/, updatingStatusText);
                })
        );
    });

    await Promise.all(fileWriteRequests);
}

function checkSnippetsForUniqueIDs(processedSnippets) {
    banner('Testing every snippet for ID uniqueness');

    let idsAllUnique = true; // assume best, until proven otherwise
    processedSnippets.values()
        .forEach(item => {
            status.add(`Testing ID of snippet ${item.relativePath}`);
            const otherMatches = processedSnippets.values().filter(anotherItem => anotherItem !== item && anotherItem.id === item.id);
            const isUnique = (otherMatches.length === 0);
            status.complete(isUnique /*succeeded*/,
                `Testing ID of snippet ${item.relativePath}`,
                isUnique ? null : [`ID "${item.id}" not unique, and matches the IDs of `].concat(otherMatches.map(item => '    ' + item.relativePath)));
            if (!isUnique) {
                idsAllUnique = false;
            }
        });

    if (!idsAllUnique) {
        throw new Error('Not all snippet IDs are unique; cannot continue');
    }
}

async function generatePlaylists(processedSnippets: Dictionary<SnippetProcessedData>) {
    banner('Generating playlists');

    let processedPublicSnippets = new Dictionary<SnippetProcessedData>();
    for (let processedSnippet of processedSnippets.values()) {
        if (processedSnippet.isPublic) {
            processedPublicSnippets.set(processedSnippet.rawUrl, processedSnippet);
        }
    }

    /* Creating playlists directory */
    status.add(`Creating \'playlists\' folder`);
    await rmRf('playlists');
    await mkDir('playlists');
    status.complete(true /*success*/, `Creating \'playlists\' folder`);

    const publicGroups = groupBy(
        processedPublicSnippets.values()
            .filter((file) => !(file == null) && file.fileName !== 'default.yaml'),
        'host');
    let publicPlaylistPromises = map(publicGroups, async (items, host) => {
        const creatingStatusText = `Creating ${host}.yaml`;
        status.add(creatingStatusText);
        items = sortBy(items, sortingCriteria) as any;

        /*
           Having sorted the items -- which may have included a number in the group name! -- remove the group number if any
           Note that by this time, the group names have already had the dash removed from them,
           and would now have <number><space> if anything at all.

            01 basics
            ==> Contains number, strip it

            basics
            ==> No number, keep as is

            hello-world-123-foo
            ==> No number, keep as is
        */
        const groupNumberRegex = /^(\d+\s)?(\w.*)$/;

        let modifiedItems = items.map(item => {
            /* Only keep select properties that are needed */
            return {
                id: item.id,
                name: item.name,
                fileName: item.fileName,
                description: item.description,
                rawUrl: item.rawUrl,
                group: item.group.replace(groupNumberRegex, '$2'),
                api_set: item.api_set,
            };
        });

        let contents = jsyaml.safeDump(modifiedItems, {
            skipInvalid: true /* skip "undefined" (e.g., for "order" on some of the snippets) */
        });

        let fileName = `playlists/${host}.yaml`;
        await writeFile(path.resolve(fileName), contents);

        status.complete(true /*success*/, creatingStatusText);
    });

    /* Creating view directory */
    status.add(`Creating \'view\' folder`);
    await rmRf('view');
    await mkDir('view');
    status.complete(true /*success*/, `Creating \'view\' folder`);

    const allGroups = groupBy(
        processedSnippets.values()
            .filter((file) => !(file == null) && file.fileName !== 'default.yaml'),
        'host');
    let allPlaylistPromises = map(allGroups, async (items, host) => {
        const creatingStatusText = `Creating ${host}.json`;
        status.add(creatingStatusText);
        items = sortBy(items, sortingCriteria) as any;

        let hostMapping = {} as { [id: string]: string };
        items.forEach(item => {
            hostMapping[item.id] = item.rawUrl;
        });

        /* Group private samples with public samples in one JSON file */
        let fileName = `view/${host}.json`;
        await writeFile(path.resolve(fileName), JSON.stringify(hostMapping, null, 2));

        status.complete(true /*success*/, creatingStatusText);
    });

    await Promise.all(publicPlaylistPromises.concat(allPlaylistPromises));
}

async function copyAndUpdatePlaylistFolders() {
    banner('Copying and updating \'playlists\' and \'view\' directories');

    /* Copying playlists directory */
    let playlistsProdFolderName = 'playlists-prod';
    status.add(`Creating \'${playlistsProdFolderName}\' folder`);
    await rmRf(playlistsProdFolderName);
    let playlistsProdFolderPath = await mkDir(playlistsProdFolderName);
    status.complete(true /*success*/, `Creating \'${playlistsProdFolderName}\' folder`);

    await fsx.copy('playlists', playlistsProdFolderName);
    let playlistsFiles = await readDir(playlistsProdFolderPath);
    (await Promise.all(playlistsFiles.map(file => updateCopiedFile(playlistsProdFolderPath, file))));

    /* Copying view directory */
    let viewProdFolderName = `view-prod`;
    status.add(`Creating \'${viewProdFolderName}\' folder`);
    await rmRf(viewProdFolderName);
    let viewProdFolderPath = await mkDir(viewProdFolderName);
    status.complete(true /*success*/, `Creating \'${viewProdFolderName}\' folder`);

    await fsx.copy('view', viewProdFolderName);
    let viewFiles = await readDir(viewProdFolderPath);
    (await Promise.all(viewFiles.map(file => updateCopiedFile(viewProdFolderPath, file))));
}

// helper for copyAndUpdatePlaylistFolders
async function updateCopiedFile(folderPath: string, filePath: string) {
    const fullPath = path.resolve(folderPath, filePath);
    let content = fsx.readFileSync(fullPath).toString().trim().replace(
        /\/OfficeDev\/office-js-snippets\/main/g,
        '/OfficeDev/office-js-snippets/prod');
    const fileUpdates = [];
    fileUpdates.push(
        Promise.resolve()
            .then(async () => {
                const updatingStatusText = `Updating copied file ${fullPath}`;
                status.add(updatingStatusText);
                await writeFile(fullPath, content);
                status.complete(true /*succeeded*/, updatingStatusText);
            })
    );
}

function handleError(error: any | any[]) {
    if (!isArray(error)) {
        error = [error];
    }

    banner('One or more errors had occurred during processing:', null, chalk.bold.red);
    (error as any[]).forEach(item => {
        console.log(chalk.red(` * ${item.message || item}`));
        if (item instanceof Error) {
            console.log(getPrintableDetails(item, 4 /*indent*/));
        }
    });

    banner('Cannot continue, closing.',
        'Note that if you were building locally, please see pending changes ' +
        'for anything that may have modified locally (and which might make you pass on a 2nd try).',
        chalk.bold.red);

    process.exit(1);
}
