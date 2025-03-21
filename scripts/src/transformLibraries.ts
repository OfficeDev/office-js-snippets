/**
 * Transform library references.
 * - Remove jquery & core-js
 * - Reference CDN for office.js types
 * - Directly reference unpkg for npm packages
 * @returns transformed libraries
 */
export function transformLibraries(data: string): string {
    function getLinkFromPackageReference(packageReference: string): string | undefined {
        const reg = /^(?<packageName>.*)@(?<packageVersion>\d+\.\d+\.\d+)\/(?<packageFile>.*)$/;
        const groups = reg.exec(packageReference)?.groups;
        if (groups === undefined) {
            return packageReference;
        }

        const { packageName, packageVersion, packageFile } = groups;

        return `https://unpkg.com/${packageName}@${packageVersion}/${packageFile}`;
    }

    const cleanLibraries = data
        .split("\n")
        .map((line) => {
            line = line.trim();

            // Empty line
            if (line === "") {
                return "";
            }

            // Comment
            if (line.startsWith("//") || line.startsWith("#")) {
                return line;
            }

            // direct reference
            if (line.startsWith("https://") || line.startsWith("http://")) {
                return line;
            }

            // office.js
            if (line === "@types/office-js") {
                return `https://appsforoffice.microsoft.com/lib/1/hosted/office.d.ts`;
            }

            // Remove packages
            const packageNamesIgnore = ["jquery", "@types/jquery", "core-js", "@types/core-js"];
            const isExcluded = packageNamesIgnore.some((packageName) =>
                line.startsWith(packageName),
            );
            if (isExcluded) {
                return undefined;
            }

            // npm reference
            const link = getLinkFromPackageReference(line);
            return link;
        })
        .filter((line) => line !== undefined) as string[];

    const cleanData = cleanLibraries.join("\n").replace(/\n\n\n/, "\n\n");
    return cleanData;
}
