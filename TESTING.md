# Testing infrastructure

This repository now includes enhanced static analysis testing for Office.js snippets.

## Test suites

### 1. TypeScript compilation tests

**File:** `tests/snippet-compiler.test.ts`

Validates that all snippet TypeScript code compiles without errors against Office.js type definitions.

- Extracts snippet code from YAML files.
- Creates compilation harness with Office.js type declarations.
- Reports compilation errors with line numbers.
- Validates API usage through TypeScript's type system.

### 2. Library URL validation tests

**File:** `tests/library-validator.test.ts`

Verifies that all library URLs referenced in snippets are reachable.

- Checks HTTP HEAD requests for all library URLs.
- Caches results to avoid repeated requests.
- Skips @types packages (npm references).
- Reports broken links with affected snippets.

## Running tests

```bash
# Run all tests
npm test

# Run specific test suites
npm run test:compile    # TypeScript compilation only
npm run test:libs       # Library URL checks only

# Run tests in watch mode
npm run test:watch

# Run tests with coverage
npm run test:coverage

# Run full validation (build + lint + tests)
npm run validate
```

## CI/CD integration

Tests run automatically on:

- Every pull request
- Pushes to `main` and `prod` branches
- Multiple Node.js versions (18.x, 20.x)

**Workflow:** `.github/workflows/ci.yml`

## Improving Type Definitions

The TypeScript compiler utility uses custom type declarations in `tests/helpers/snippet-compiler.ts`. As compilation tests identify missing APIs, these declarations can be expanded.

**Common issues:**

- Missing Office.js API methods
- Incomplete type definitions for newer APIs
- Namespace visibility issues

To add support for a new API:

1. Identify the error from test output
2. Add the API to the appropriate class/interface in `snippet-compiler.ts`
3. Re-run tests to verify

## Future Enhancements

### Phase 2: Office Automation Smoke Tests

- E2E testing of critical snippets in real Office applications.
- Requires self-hosted GitHub Actions runner with Windows + Office.
- Would catch actual API regressions vs type/compatibility issues.

## Troubleshooting

### Tests timing out

The full compilation test suite (335 snippets) can take 2-3 minutes. This is normal.

### False positives

Some snippets may fail compilation due to incomplete type definitions, not actual code errors. Check the error message - if it's about missing type definitions, the types can be added to `snippet-compiler.ts`.

### Library URL failures

Transient network issues may cause URL validation failures. The test includes retry logic, but occasional failures from CDN downtime are possible.

## Contributing

When adding new snippets:

1. Run `npm run validate` before committing.
2. Fix any test failures.
3. If compilation fails due to missing type definitions, either:
   - Add the types to `snippet-compiler.ts`, or
   - Document the issue and skip that test temporarily.

The CI pipeline will catch issues automatically on pull requests.
