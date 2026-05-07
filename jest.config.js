// Disable Application Insights statsbeat telemetry before any test worker starts.
// Without this, the Azure Monitor exporter starts a background timer that fires
// a dynamic import() inside Node's VM context. Node 24 disallows this without
// --experimental-vm-modules, crashing the process after test teardown with
// ERR_VM_DYNAMIC_IMPORT_CALLBACK_MISSING_FLAG.
process.env['APPLICATION_INSIGHTS_NO_STATSBEAT'] = '1';

module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'node',
  roots: ['<rootDir>/tests'],
  testMatch: ['**/*.test.ts'],
  testPathIgnorePatterns: ['/node_modules/', 'quick-test.test.ts'],
  moduleFileExtensions: ['ts', 'tsx', 'js', 'jsx', 'json'],
  collectCoverageFrom: [
    'tests/**/*.ts',
    '!tests/**/*.test.ts',
    '!tests/**/*.d.ts'
  ],
  coverageDirectory: 'coverage',
  verbose: true,
  testTimeout: 30000, // 30 seconds for URL validation tests
  transform: {
    '^.+\\.ts$': ['ts-jest', {
      tsconfig: 'tsconfig.test.json'
    }]
  }
};
