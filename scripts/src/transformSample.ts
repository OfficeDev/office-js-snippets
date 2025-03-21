import { RawSample } from "./RawSample";
import { transformCss } from "./transformCss";
import { transformHtml } from "./transformHtml";
import { transformLibraries } from "./transformLibraries";
import { transformTypeScript } from "./transformTypeScript";

export function transformRawSample(id: string, rawSample: RawSample): RawSample {
    const typescriptRaw = rawSample?.script?.content;
    const htmlRaw = rawSample?.template?.content;
    const cssRaw = rawSample?.style?.content;
    const librariesRaw = rawSample?.libraries;

    if ([typescriptRaw, htmlRaw, cssRaw, librariesRaw].some((content) => content === undefined)) {
        console.log(`ERROR: Empty content [${rawSample.name}] ${id}`);
        // happens for custom functions
        return rawSample;
    }

    const typescriptContent = transformTypeScript(typescriptRaw).trim();
    const htmlContent = transformHtml(htmlRaw).trim();
    const cssContent = transformCss(cssRaw).trim();
    const librariesContent = transformLibraries(librariesRaw).trim();

    // Update the raw sample with the transformed content
    rawSample.script.content = typescriptContent;
    rawSample.template.content = htmlContent;
    rawSample.style.content = cssContent;
    rawSample.libraries = librariesContent;

    return rawSample;
}
