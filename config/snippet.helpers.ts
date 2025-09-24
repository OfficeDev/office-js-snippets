/////////////////////////////////////////////////////////////////////////////////////////
/// NOTE: This file is a (partial) copy of "src\client\app\helpers\snippet.helper.ts" ///
///       of the root project. Please be sure that any changes that you               ///
///       make here (or vice-versa) are applied to both projects.                     ///
/////////////////////////////////////////////////////////////////////////////////////////

import * as jsyaml from 'js-yaml';
import { forIn } from 'lodash';

export enum SnippetFieldType {
    /** PUBLIC = Store internally, and also include in copy-to-clipboard */
    PUBLIC = 1 << 0,

    /** INTERNAL = Necessary to store, but not copy out */
    INTERNAL = 1 << 1,

    /** TRANSIENT = Only useful at runtime, needn't be stored at all */
    TRANSIENT = 1 << 2
}

const snippetFields: { [key: string]: SnippetFieldType } = {
    /* ITemplate base class */
    id: SnippetFieldType.INTERNAL,
    gist: SnippetFieldType.INTERNAL,
    name: SnippetFieldType.PUBLIC,
    description: SnippetFieldType.PUBLIC,
    // author: export-only, always want to generate on the fly, so skip altogether
    host: SnippetFieldType.PUBLIC,
    // api_set: export-only, always want to generate on the fly, so skip altogether
    platform: SnippetFieldType.TRANSIENT,
    origin: SnippetFieldType.TRANSIENT,
    created_at: SnippetFieldType.INTERNAL,
    modified_at: SnippetFieldType.INTERNAL,

    /* ISnippet */
    script: SnippetFieldType.PUBLIC,
    template: SnippetFieldType.PUBLIC,
    style: SnippetFieldType.PUBLIC,
    libraries: SnippetFieldType.PUBLIC
};

export const snippetFieldSortingOrder: { [key: string]: number } = {
    /* Sample-exported fields */
    order: 1,
    id: 2,

    /* ITemplate base class */
    name: 11,
    description: 12,
    author: 13,
    host: 14,
    api_set: 15,

    /* ISnippet */
    script: 110,
    template: 111,
    style: 112,
    libraries: 113,

    /* And within scripts / templates / styles, content should always be before language */
    content: 1000,
    language: 1001
};

function scrubCarriageReturns(snippet: ISnippet) {
    removeCarriageReturns(snippet, 'template');
    removeCarriageReturns(snippet, 'script');
    removeCarriageReturns(snippet, 'style');
    removeCarriageReturns(snippet, 'libraries');

    function removeCarriageReturns(snippet: ISnippet, field: 'template' | 'script' | 'style' | 'libraries') {
        if (!snippet[field]) {
            return;
        }

        if (field === 'libraries') {
            snippet.libraries = removeCarriageReturnsHelper(snippet.libraries);
        } else {
            snippet[field].content = removeCarriageReturnsHelper(snippet[field].content);
        }

        function removeCarriageReturnsHelper(text) {
            return text
                .split('\n')
                .map(line => line.replace(/\r/, ''))
                .join('\n');
        }
    }
}
/** Returns a shallow copy of the snippet, filtered to only keep a particular set of fields */
export function getScrubbedSnippet(snippet: ISnippet, keep: SnippetFieldType): ISnippet {
    let copy = {};
    forIn(snippetFields, (fieldType, fieldName) => {
        if (fieldType & keep && snippet[fieldName] !== undefined) {
            copy[fieldName] = snippet[fieldName];
        }
    });

    return copy as ISnippet;
}

export function getShareableYaml(rawSnippet: ISnippet, additionalFields: ISnippet) {
    const snippet = { ...getScrubbedSnippet(rawSnippet, SnippetFieldType.PUBLIC), ...additionalFields };
    scrubCarriageReturns(snippet);

    return jsyaml.dump(snippet, {
        indent: 4,
        lineWidth: -1,
        sortKeys: <any>((a, b) => snippetFieldSortingOrder[a] - snippetFieldSortingOrder[b]),
        skipInvalid: true
    });
}
