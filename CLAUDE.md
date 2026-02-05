# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

This repository contains a collection of Office JavaScript API code snippets for Script Lab. Snippets are YAML files containing TypeScript/JavaScript code samples demonstrating Office.js APIs for Excel, Word, PowerPoint, Outlook, OneNote, and Project.

## Build Commands

```bash
# Main build command - validates, compiles, and lints all snippets
npm start

# Individual steps (npm start runs all three):
npm run tsc      # Compile TypeScript in config/ directory
npm run build    # Process and validate all snippets
npm run lint     # Run TSLint on config/ files
```

The build process:
1. Validates all `.yaml` snippet files in `samples/` and `private-samples/`
2. Auto-corrects common issues (IDs, formatting, tabs → spaces)
3. Generates playlists in `playlists/` and `playlists-prod/`
4. Generates view mappings in `view/` and `view-prod/`
5. Updates snippet files with corrected content

## Architecture

### Repository Structure

- **`samples/{host}/{group}/`** - Public snippets organized by Office host (excel, word, powerpoint, outlook, onenote, project) and feature group
- **`private-samples/{host}/{group}/`** - Internal/unpublished snippets
- **`config/`** - TypeScript build scripts and validation logic
- **`playlists/`** - Generated YAML playlist files (main branch URLs)
- **`playlists-prod/`** - Generated YAML playlist files (prod branch URLs)
- **`view/`** - Generated JSON mappings of snippet IDs to URLs (includes private)
- **`view-prod/`** - Generated JSON mappings (prod branch URLs)
- **`snippet-extractor-metadata/`** - Documentation extraction metadata
- **`snippet-extractor-output/`** - Documentation extraction output

### Snippet YAML Structure

Each snippet is a YAML file with:
- **Metadata**: `id`, `name`, `description`, `author`, `host`, `api_set`, `order` (optional)
- **`script`**: TypeScript/JavaScript code with `content` and `language`
- **`template`**: HTML markup with `content` and `language`
- **`style`**: CSS styling with `content` and `language`
- **`libraries`**: Newline-separated list of script/CSS URLs

Groups prefixed with numbers (e.g., `01-basics`, `10-chart`) control ordering in Script Lab.

### Build System (config/)

The build process is orchestrated by `config/build.ts`:

1. **Snippet Processing** (`processSnippets()`):
   - Scans `samples/` and `private-samples/` for `.yaml` files
   - Validates each snippet (see validations below)
   - Auto-generates/corrects snippet IDs based on file path
   - Converts tabs to spaces, normalizes formatting

2. **Validations**:
   - **ID uniqueness**: All snippet IDs must be globally unique
   - **Naming**: File names must be kebab-case with `.yaml` extension
   - **Host matching**: Snippet `host` field must match directory path
   - **API sets**: Excel, Word, OneNote require `api_set` specification (e.g., `ExcelApi: '1.5'`)
   - **Office.js references**: Office snippets must reference official office.js and office.d.ts URLs
   - **Library versions**: All libraries must specify versions (no unpkg, use versioned CDNs)
   - **Fabric UI**: If using `office-ui-fabric-core`, must also include `office-ui-fabric-js` components CSS
   - **TypeScript declarations**: Auto-converts `dt~` notation to `@types/` format

3. **Playlist Generation** (`generatePlaylists()`):
   - Creates per-host YAML files (`playlists/{host}.yaml`) with public snippets sorted by group/order
   - Creates per-host JSON files (`view/{host}.json`) mapping snippet IDs to raw GitHub URLs
   - Generates both main-branch and prod-branch versions

### Key Build Script Modules

- **`config/build.ts`** - Main orchestration logic
- **`config/snippet.helpers.ts`** - YAML serialization utilities
- **`config/libraries.processor.ts`** - Library reference validation
- **`config/build.documentation.ts`** - Documentation extraction for reference docs
- **`config/helpers.ts`** - File I/O and utility functions
- **`config/status.ts`** - Console output formatting

## Snippet Development Workflow

### Adding a New Snippet

1. Create snippet in Script Lab, copy YAML via Share → Copy to Clipboard
2. Set `api_set` to the highest API version used (e.g., `ExcelApi: '1.5'`)
3. Save as `kebab-case-name.yaml` in appropriate `samples/{host}/{group}/` folder
4. If group has ordered snippets (check for `order` property in existing files), add `order` field
5. Run `npm start` - the build will auto-generate the snippet `id` and validate
6. Fix any validation errors and re-run `npm start` until successful
7. Commit both the new `.yaml` file and the modified `playlists/{host}.yaml`

### Snippet Style Guidelines

All snippets should follow this pattern:

```typescript
document.getElementById("run").addEventListener("click", () => tryCatch(run));

async function run() {
    await Excel.run(async (context) => {
        // API calls here
        await context.sync();
    });
}

async function tryCatch(callback) {
    try {
        await callback();
    } catch (error) {
        console.error(error);
    }
}
```

**Style rules**:
- Standard TypeScript indentation (4 spaces, no tabs)
- Double quotes for strings (enforced by tslint on build scripts)
- Semicolons required
- Button click handlers invoke `tryCatch()` wrapper
- HTML IDs use `all-lower-case-and-hyphenated`
- Library references must include version numbers (e.g., `jquery@3.1.1`)
- Keep snippets small (few hundred lines max)
- Avoid over-engineering beyond what the sample demonstrates

## Common Validation Errors

- **"name has upper-case letters or other disallowed characters"** - This refers to the file name, not the `name` property. Use kebab-case.
- **"Snippet host is different than the directory path host"** - The `host` field must match the directory (e.g., file in `samples/excel/` must have `host: EXCEL`)
- **"No API set specified"** - Excel, Word, and OneNote snippets require the `api_set` field
- **"Office.js reference does not match canonical form"** - Use `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` (or `/lib/beta/` for preview APIs)
- **"ID not unique"** - Another snippet already has this ID. The build will suggest a unique ID based on file path.

## Testing and Debugging

- **Test in Script Lab**: Import via Script Lab's Samples gallery or paste YAML directly
- **Debug build scripts**:
  - Build scripts are in `config/` written in TypeScript
  - Run `npm run tsc` to recompile after changes
  - In VS Code, use F5 to attach debugger (must run `npm run tsc` first)
  - Build script entry point: `config/build.ts`
