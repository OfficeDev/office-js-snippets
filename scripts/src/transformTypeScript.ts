/**
 * Transform TypeScript code.
 * - remove JQuery handlers
 * - Add Office on ready.
 */
export function transformTypeScript(data: string): string {
    // remove jquery
    // $("#id").on("click", () => tryCatch(handler));`;
    const jqueryReg = /^\$\("#(?<id>.*)"\)\.on\("click", \(\) => tryCatch\((?<handler>.*)\)\);$/;

    // Outlook specific
    // $("#id").on("click", handler);
    const jqueryAlt = /^\$\("#(?<id>.*)"\)\.on\("click", (?<handler>.*)\);$/;
    // $("#id").click(handler);
    const jqueryAlt2 = /^\$\("#(?<id>.*)"\)\.click\((?<handler>.*)\);$/;

    const cleanData = data
        .split("\n")
        .map((line) => {
            const trimLine = line.trim();

            if (trimLine.startsWith("$")) {
                // JQuery
                const match = jqueryReg.exec(trimLine);
                if (match !== null) {
                    const groups = match?.groups;
                    if (groups) {
                        const { id, handler } = groups;
                        return `document.getElementById("${id}").addEventListener("click", () => tryCatch(${handler}));`;
                    }
                }

                const matchAlt = jqueryAlt.exec(trimLine) || jqueryAlt2.exec(trimLine);
                if (matchAlt !== null) {
                    const groups = matchAlt?.groups;
                    if (groups) {
                        const { id, handler } = groups;
                        return `document.getElementById("${id}").addEventListener("click", ${handler});`;
                    }
                }
            }

            return line;
        })
        .join("\n");

    const code = cleanData;

    return code;
}
