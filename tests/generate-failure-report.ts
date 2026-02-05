import { getAllSnippets, hasTypeScriptCode } from './helpers/test-helpers';
import { compileSnippet } from './helpers/snippet-compiler';
import * as fs from 'fs';
import * as path from 'path';

interface FailureReport {
  snippet: string;
  host: string;
  errors: Array<{
    message: string;
    line?: number;
    column?: number;
  }>;
}

async function generateFailureReport() {
  console.log('Generating compilation failure report...\n');

  const snippets = getAllSnippets().filter(hasTypeScriptCode);
  const failures: FailureReport[] = [];

  let passed = 0;
  let failed = 0;

  for (const snippet of snippets) {
    const result = compileSnippet(snippet);

    if (!result.success) {
      failed++;
      failures.push({
        snippet: snippet.relativePath,
        host: snippet.host || 'UNKNOWN',
        errors: result.errors
      });
    } else {
      passed++;
    }

    // Progress indicator
    if ((passed + failed) % 50 === 0) {
      console.log(`  Processed ${passed + failed}/${snippets.length} snippets...`);
    }
  }

  console.log(`\nCompleted: ${passed} passed, ${failed} failed\n`);

  // Generate report
  const reportLines: string[] = [];
  reportLines.push('# TypeScript Compilation Failure Report');
  reportLines.push('');
  reportLines.push(`**Generated:** ${new Date().toISOString()}`);
  reportLines.push(`**Total Snippets:** ${snippets.length}`);
  reportLines.push(`**Passing:** ${passed} (${Math.round(passed / snippets.length * 100)}%)`);
  reportLines.push(`**Failing:** ${failed} (${Math.round(failed / snippets.length * 100)}%)`);
  reportLines.push('');
  reportLines.push('---');
  reportLines.push('');

  if (failures.length === 0) {
    reportLines.push('## ✅ All snippets compile successfully!');
  } else {
    // Group by host
    const byHost = new Map<string, FailureReport[]>();
    failures.forEach(f => {
      const host = f.host.toUpperCase();
      if (!byHost.has(host)) {
        byHost.set(host, []);
      }
      byHost.get(host)!.push(f);
    });

    // Summary by host
    reportLines.push('## Summary by Host');
    reportLines.push('');
    byHost.forEach((failures, host) => {
      reportLines.push(`- **${host}**: ${failures.length} failures`);
    });
    reportLines.push('');
    reportLines.push('---');
    reportLines.push('');

    // Collect all unique error types
    const errorTypes = new Map<string, number>();
    failures.forEach(f => {
      f.errors.forEach(e => {
        // Extract error pattern (e.g., "Cannot find name 'X'" -> "Cannot find name")
        const pattern = e.message
          .replace(/'[^']+'/g, "'...'")
          .replace(/"[^"]+"/g, '"..."')
          .replace(/\b\d+\b/g, 'N');
        errorTypes.set(pattern, (errorTypes.get(pattern) || 0) + 1);
      });
    });

    // Most common errors
    reportLines.push('## Most Common Error Types');
    reportLines.push('');
    const sortedErrors = Array.from(errorTypes.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 20);

    sortedErrors.forEach(([pattern, count]) => {
      reportLines.push(`- (${count}x) ${pattern}`);
    });
    reportLines.push('');
    reportLines.push('---');
    reportLines.push('');

    // Detailed failures by host
    reportLines.push('## Detailed Failures');
    reportLines.push('');

    byHost.forEach((failures, host) => {
      reportLines.push(`### ${host} (${failures.length} failures)`);
      reportLines.push('');

      failures.forEach(failure => {
        reportLines.push(`#### \`${failure.snippet}\``);
        reportLines.push('');

        failure.errors.forEach((error, idx) => {
          const location = error.line ? ` (line ${error.line}:${error.column})` : '';
          reportLines.push(`${idx + 1}. **Error${location}:**`);
          reportLines.push(`   \`\`\``);
          reportLines.push(`   ${error.message}`);
          reportLines.push(`   \`\`\``);
          reportLines.push('');
        });
      });

      reportLines.push('');
    });
  }

  // Write to file
  const reportPath = path.resolve(__dirname, '../COMPILATION-FAILURES.md');
  fs.writeFileSync(reportPath, reportLines.join('\n'));

  console.log(`Report written to: ${reportPath}`);
  console.log('');

  // Also write CSV for easy parsing
  const csvLines: string[] = [];
  csvLines.push('Snippet,Host,Error Message,Line,Column');
  failures.forEach(f => {
    f.errors.forEach(e => {
      const snippet = f.snippet.replace(/,/g, ';');
      const message = e.message.replace(/,/g, ';').replace(/\n/g, ' ');
      csvLines.push(`"${snippet}","${f.host}","${message}",${e.line || ''},${e.column || ''}`);
    });
  });

  const csvPath = path.resolve(__dirname, '../COMPILATION-FAILURES.csv');
  fs.writeFileSync(csvPath, csvLines.join('\n'));

  console.log(`CSV report written to: ${csvPath}`);
  console.log('');

  // Exit with non-zero code if there are failures
  process.exit(failures.length > 0 ? 1 : 0);
}

generateFailureReport().catch(error => {
  console.error('Error generating report:', error);
  process.exit(1);
});
