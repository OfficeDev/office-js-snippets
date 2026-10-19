import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { loadNycConfig } from '@istanbuljs/load-nyc-config';

describe('Coverage configuration with the js-yaml override', () => {
  let directory: string;

  beforeEach(() => {
    directory = fs.mkdtempSync(path.join(os.tmpdir(), 'snippet-coverage-config-'));
  });

  afterEach(() => {
    fs.rmSync(directory, { recursive: true, force: true });
  });

  test('loads YAML, inherited settings, and camel-cased options', async () => {
    fs.writeFileSync(path.join(directory, 'base.yaml'), 'extension:\n  - .ts\nreporter:\n  - text\n');
    fs.writeFileSync(
      path.join(directory, '.nycrc.yaml'),
      'extends: ./base.yaml\ncheck-coverage: true\nbranches: 80\n'
    );

    const config = await loadNycConfig({ cwd: directory, nycrcPath: '.nycrc.yaml' });

    expect(config).toMatchObject({
      cwd: directory,
      extension: ['.ts'],
      reporter: ['text'],
      checkCoverage: true,
      branches: 80,
    });
  });

  test('rejects malformed YAML instead of ignoring configuration errors', async () => {
    fs.writeFileSync(path.join(directory, '.nycrc.yaml'), 'extension: [\n');

    await expect(loadNycConfig({ cwd: directory, nycrcPath: '.nycrc.yaml' })).rejects.toThrow();
  });
});
