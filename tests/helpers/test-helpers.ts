import * as path from 'path';
import * as fs from 'fs';
import * as jsyaml from 'js-yaml';
import { getFiles, SnippetFileInput } from '../../config/helpers';

const SAMPLES_DIR = path.resolve(__dirname, '../../samples');
const PRIVATE_SAMPLES_DIR = path.resolve(__dirname, '../../private-samples');

export interface TestSnippet {
  id: string;
  name: string;
  description: string;
  host: string;
  api_set: { [index: string]: string | number };
  script?: { content: string; language: string };
  template?: { content: string; language: string };
  style?: { content: string; language: string };
  libraries?: string;
  relativePath: string;
  fullPath: string;
}

/**
 * Get all snippet files from samples and private-samples directories
 */
export function getAllSnippetFiles(): SnippetFileInput[] {
  let files: SnippetFileInput[] = [];

  // Get public samples
  if (fs.existsSync(SAMPLES_DIR)) {
    files = files.concat(getFiles(SAMPLES_DIR));
  }

  // Get private samples if directory exists
  if (fs.existsSync(PRIVATE_SAMPLES_DIR)) {
    files = files.concat(getFiles(PRIVATE_SAMPLES_DIR));
  }

  return files;
}

/**
 * Load and parse a snippet YAML file
 */
export function loadSnippet(file: SnippetFileInput): TestSnippet {
  const content = fs.readFileSync(file.fullPath, 'utf8');
  const snippet = jsyaml.load(content) as any;

  return {
    ...snippet,
    relativePath: file.relativePath,
    fullPath: file.fullPath
  };
}

/**
 * Get all snippets as test data
 */
export function getAllSnippets(): TestSnippet[] {
  const files = getAllSnippetFiles();
  return files.map(file => loadSnippet(file));
}

/**
 * Filter snippets by host
 */
export function getSnippetsByHost(host: string): TestSnippet[] {
  return getAllSnippets().filter(s =>
    s.host && s.host.toUpperCase() === host.toUpperCase()
  );
}

/**
 * Check if a snippet has TypeScript code
 */
export function hasTypeScriptCode(snippet: TestSnippet): boolean {
  return snippet.script &&
         snippet.script.content &&
         snippet.script.content.trim().length > 0;
}

/**
 * Extract Office.js library URLs from libraries string
 */
export function getOfficeJsLibraries(snippet: TestSnippet): string[] {
  if (!snippet.libraries) {
    return [];
  }

  return snippet.libraries
    .split('\n')
    .map(lib => lib.trim())
    .filter(lib => lib.length > 0);
}

/**
 * Get the API set version for a snippet's host
 */
export function getApiSetVersion(snippet: TestSnippet): number | undefined {
  if (!snippet.api_set || !snippet.host) {
    return undefined;
  }

  const apiSetKey = `${snippet.host}Api`;
  const version = snippet.api_set[apiSetKey];

  if (typeof version === 'string') {
    return parseFloat(version);
  }

  return version;
}

/**
 * Load a snippet by relative path
 */
export function loadSnippetByPath(relativePath: string): TestSnippet {
  const fullPath = path.resolve(__dirname, '../../', relativePath);

  // Extract metadata from path (e.g., "samples/excel/01-basics/basic-api-call.yaml")
  const parts = relativePath.split(path.sep);
  const file_name = parts[parts.length - 1];
  const isPublic = parts[0] === 'samples';
  const host = parts.length > 1 ? parts[1] : '';
  const group = parts.length > 2 ? parts[2] : '';

  const file: SnippetFileInput = {
    file_name,
    relativePath,
    fullPath,
    host,
    group,
    isPublic,
  };

  return loadSnippet(file);
}

/**
 * Check if a snippet uses preview Office.js APIs
 */
export function usesPreviewApis(snippet: TestSnippet): boolean {
  // Check if libraries field contains office-js-preview
  if (snippet.libraries && snippet.libraries.includes('office-js-preview')) {
    return true;
  }

  // Check if api_set contains 'preview' value
  if (snippet.api_set) {
    for (const key in snippet.api_set) {
      if (snippet.api_set[key] === 'preview') {
        return true;
      }
    }
  }

  return false;
}

/**
 * Detect external libraries used in a snippet
 * Returns array of @types package names to include
 */
export function getExternalLibraryTypes(snippet: TestSnippet): string[] {
  const types: string[] = [];

  if (!snippet.libraries && !snippet.script?.content && !snippet.api_set) {
    return types;
  }

  const libraries = snippet.libraries || '';
  const code = snippet.script?.content || '';

  // Check for Lodash
  if (libraries.includes('lodash') || code.match(/\b_\./)) {
    types.push('lodash');
  }

  // Check for Custom Functions Runtime
  if (snippet.api_set && snippet.api_set['CustomFunctionsRuntime']) {
    types.push('custom-functions-runtime');
  }

  return types;
}
