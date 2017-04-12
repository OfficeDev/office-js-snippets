/////////////////////////////////////////////////////////////////////////////////////////////////////
/// NOTE: This file is an identical copy from the main Office JavaScript API Snippets project.                ///
///       Please be sure that any changes that you make here are also copied to there.            ///
///       See "src/server/core/libraries.processor.ts" in https://github.com/OfficeDev/project-bornholm ///
/////////////////////////////////////////////////////////////////////////////////////////////////////


export function processLibraries(snippet: ISnippet) {
    let linkReferences: string[] = [];
    let scriptReferences: string[] = [];
    let officeJS: string = null;

    snippet.libraries.split('\n').forEach(processLibrary);

    return { linkReferences, scriptReferences, officeJS };


    function processLibrary(text: string) {
        if (text == null || text.trim() === '') {
            return null;
        }

        text = text.trim();

        let isNotScriptOrStyle =
            /^#.*|^\/\/.*|^\/\*.*|.*\*\/$.*/im.test(text) ||
            /^@types/.test(text) ||
            /^dt~/.test(text) ||
            /\.d\.ts$/i.test(text);

        if (isNotScriptOrStyle) {
            return null;
        }

        let resolvedUrlPath = (/^https?:\/\/|^ftp? :\/\//i.test(text)) ? text : `https://unpkg.com/${text}`;

        if (/\.css$/i.test(resolvedUrlPath)) {
            return linkReferences.push(resolvedUrlPath);
        }

        if (/\.ts$|\.js$/i.test(resolvedUrlPath)) {
            /*
            * Don't add Office.js to the rest of the script references --
            * it is special because of how it needs to be *outside* of the iframe,
            * whereas the rest of the script references need to be inside the iframe.
            */
            if (/(?:office|office.debug).js$/.test(resolvedUrlPath.toLowerCase())) {
                officeJS = resolvedUrlPath;
                return null;
            }

            return scriptReferences.push(resolvedUrlPath);
        }

        return scriptReferences.push(resolvedUrlPath);
    }
}
