import * as ts from 'typescript';
import * as path from 'path';
import * as fs from 'fs';
import { TestSnippet } from './test-helpers';

// Load API version catalog
const apiVersionsPath = path.join(__dirname, '../data/api-versions.json');
const apiVersions: { [host: string]: { [api: string]: number } } = JSON.parse(
  fs.readFileSync(apiVersionsPath, 'utf8')
);

export interface ApiUsage {
  api: string;
  line?: number;
  column?: number;
}

export interface ApiVersionCheck {
  api: string;
  requiredVersion: number;
  declaredVersion: number;
  compatible: boolean;
  line?: number;
  column?: number;
}

/**
 * Parse Office.js API calls from TypeScript code
 */
export function parseOfficeApiCalls(code: string, host: string): ApiUsage[] {
  const sourceFile = ts.createSourceFile(
    'snippet.ts',
    code,
    ts.ScriptTarget.Latest,
    true
  );

  const apiCalls: ApiUsage[] = [];
  const hostNamespaces = [host, 'Office'];

  function visit(node: ts.Node) {
    // Look for property access expressions like Excel.Range.format.fill.color
    if (ts.isPropertyAccessExpression(node)) {
      const apiPath = getApiPath(node);

      // Check if this is an Office.js API call
      if (apiPath && hostNamespaces.some(ns => apiPath.startsWith(`${ns}.`))) {
        const { line, character } = sourceFile.getLineAndCharacterOfPosition(node.getStart());

        apiCalls.push({
          api: apiPath,
          line: line + 1,
          column: character + 1
        });
      }
    }

    ts.forEachChild(node, visit);
  }

  visit(sourceFile);

  return apiCalls;
}

/**
 * Get the full API path from a property access expression
 */
function getApiPath(node: ts.PropertyAccessExpression): string {
  const parts: string[] = [];

  function collectParts(n: ts.Node) {
    if (ts.isPropertyAccessExpression(n)) {
      collectParts(n.expression);
      parts.push(n.name.text);
    } else if (ts.isIdentifier(n)) {
      parts.push(n.text);
    }
  }

  collectParts(node);
  return parts.join('.');
}

/**
 * Check API version compatibility for a snippet
 */
export function checkApiVersions(snippet: TestSnippet): ApiVersionCheck[] {
  if (!snippet.script || !snippet.script.content) {
    return [];
  }

  const host = snippet.host ? snippet.host.toUpperCase() : 'UNKNOWN';
  const declaredVersion = getDeclaredApiVersion(snippet);

  if (declaredVersion === undefined) {
    return []; // No API set declared, skip validation
  }

  const apiCalls = parseOfficeApiCalls(snippet.script.content, host);
  const checks: ApiVersionCheck[] = [];

  for (const apiCall of apiCalls) {
    const normalizedApi = normalizeApiPath(apiCall.api, host);
    const requiredVersion = getApiVersion(normalizedApi, host);

    if (requiredVersion !== undefined) {
      checks.push({
        api: apiCall.api,
        requiredVersion,
        declaredVersion,
        compatible: requiredVersion <= declaredVersion,
        line: apiCall.line,
        column: apiCall.column
      });
    }
  }

  return checks;
}

/**
 * Get the declared API version from a snippet
 */
function getDeclaredApiVersion(snippet: TestSnippet): number | undefined {
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
 * Normalize API path by removing host namespace
 * Excel.Range.format.fill.color -> Range.format.fill.color
 */
function normalizeApiPath(apiPath: string, host: string): string {
  // Remove host namespace prefix
  const hostPrefix = `${host}.`;
  if (apiPath.startsWith(hostPrefix)) {
    return apiPath.substring(hostPrefix.length);
  }

  // Remove Office namespace prefix
  const officeHostPrefix = `Office.${host}.`;
  if (apiPath.startsWith(officeHostPrefix)) {
    return apiPath.substring(officeHostPrefix.length);
  }

  return apiPath;
}

/**
 * Get the minimum required API version for an API path
 */
function getApiVersion(apiPath: string, host: string): number | undefined {
  const hostApis = apiVersions[host];
  if (!hostApis) {
    return undefined;
  }

  // Try exact match first
  if (hostApis[apiPath] !== undefined) {
    return hostApis[apiPath];
  }

  // Try partial matches (e.g., Range.format.fill.color -> Range.format)
  const parts = apiPath.split('.');
  for (let i = parts.length - 1; i > 0; i--) {
    const partialPath = parts.slice(0, i).join('.');
    if (hostApis[partialPath] !== undefined) {
      return hostApis[partialPath];
    }
  }

  return undefined;
}

/**
 * Get all incompatible APIs used in a snippet
 */
export function getIncompatibleApis(snippet: TestSnippet): ApiVersionCheck[] {
  const checks = checkApiVersions(snippet);
  return checks.filter(check => !check.compatible);
}
