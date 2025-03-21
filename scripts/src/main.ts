//
// main entry point for the application
//

import * as fs from "fs";
import { readFileText } from "./utils/readFileText";
import { parseRawPlaylist } from "./parseRawPlaylist";
import { parseRawSample } from "./parseRawSample";
import { transformRawSample } from "./transformSample";
import yaml from "yaml";

console.log("Start edit sample yaml");

// (1) Read the playlist YAML file from sample
// (2) Read each sample YAML file
// (3) Pase the YAML file
// (4) Transform the YAML file
// (5) Write the YAML file over the original file

const sampleDirectory = "../samples";
const playlistDirectory = "../playlists-prod";

//
// Get sample files
//

const playlistFiles = fs.readdirSync(playlistDirectory);
console.log(`Playlist files:
  ${playlistFiles.join("\n  ")}`);

const playlists = playlistFiles.map((file) => {
    const filePath = `${playlistDirectory}/${file}`;
    const fileText = readFileText(filePath);
    const playlist = parseRawPlaylist(fileText);
    return playlist;
});

const playlistSamplePaths = playlists
    .map((playlist) => {
        const sampleFilePaths = playlist.map((item) => {
            const { rawUrl } = item;

            // flip raw url to the file path
            //  https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/project/basics/basic-common-api-call.yaml
            const filePath = rawUrl.replace(
                "https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/",
                "../",
            );
            return filePath;
        });

        return sampleFilePaths;
    })
    .flat();

const defaultSamplePaths = fs.readdirSync(sampleDirectory).map((file) => {
    const filePath = `${sampleDirectory}/${file}/default.yaml`;
    return filePath;
});

const samplePaths = [...defaultSamplePaths, ...playlistSamplePaths];

console.log(`Sample files:
  ${samplePaths.join("\n  ")}`);

// Check that all of the sample files exist
const checkSampleFiles = samplePaths.map((path) => {
    const present = fs.existsSync(path);
    return { present, path };
});

const missingSampleFiles = checkSampleFiles
    .filter(({ present }) => !present)
    .map(({ path }) => path);

if (missingSampleFiles.length > 0) {
    console.log("=".repeat(80));
    console.error(`Missing sample files:
      ${missingSampleFiles.join("\n")}`);
}

//
// Transform each sample file
//
console.log("=".repeat(80));
console.log("Transforming sample files...");
const transformSampleSuccess = samplePaths.map((path) => {
    console.log(`${path}`);
    let success = true;
    try {
        const fileText = readFileText(path);
        const sample = parseRawSample(fileText);
        const transformedSample = transformRawSample(path, sample);

        const transformedSampleYaml = yaml.stringify(transformedSample, {
            indent: 4,
            singleQuote: true,
        });

        fs.writeFileSync(path, transformedSampleYaml);
        console.log(`success`);
    } catch (error) {
        console.error(`ERROR\n${error}`);
        success = false;
    }

    return {
        path,
        success,
    };
});

const transformSampleErrors = transformSampleSuccess.filter(({ success }) => !success);
if (transformSampleErrors.length > 0) {
    console.log("=".repeat(80));
    console.log(`Error: Transforming sample files:
  ${transformSampleErrors.map((x) => x.path).join("\n  ")}`);
}
