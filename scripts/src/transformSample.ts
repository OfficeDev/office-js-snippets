import { RawSample } from "./RawSample";
import { transformCss } from "./transformCss";
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

    const typescriptContent = transformTypeScript(typescriptRaw);
    const htmlContent = htmlRaw;
    const cssContent = transformCss(cssRaw);
    const librariesContent = transformLibraries(librariesRaw);

    // Update the raw sample with the transformed content
    rawSample.script.content = typescriptContent;
    rawSample.template.content = htmlContent;
    rawSample.style.content = cssContent;
    rawSample.libraries = librariesContent;

    return rawSample;
}
