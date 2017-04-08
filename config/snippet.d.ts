interface ITemplate {
    ///////////////////////////////////////////////////////////////////////////////////////////////////
    // NOTE: if you add or remove any top-level fields from this list, be sure
    // to update "snippetFields" and "getSnippetDefaults" in "src\client\app\helpers\snippet.helper.ts"
    ///////////////////////////////////////////////////////////////////////////////////////////////////
    id?: string;
    gist?: string;
    name?: string;
    description?: string;
    author?: string;
    host: string;
    api_set: {
        [index: string]: number
    },
    platform: string;
    origin: string;
    created_at: number;
    modified_at: number;
}

interface ISnippet extends ITemplate {
    ///////////////////////////////////////////////////////////////////////////////////////////////////
    // NOTE: if you add or remove any top-level fields from this list, be sure
    // to update "snippetFields" and "getSnippetDefaults" in "src\client\app\helpers\snippet.helper.ts"
    ///////////////////////////////////////////////////////////////////////////////////////////////////
    script?: {
        content: string;
        language: string;
    };
    template?: {
        content: string;
        language: string;
    };
    style?: {
        content: string;
        language: string;
    };
    libraries?: string;
}

interface ILibraryDefinition {
    label?: string;
    typings?: string | string[];
    value?: string | string[];
    description?: string
}