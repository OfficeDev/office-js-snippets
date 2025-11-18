# TypeScript Style Analysis Script
# Analyzes all .yaml files in samples folder for common style issues

$samplesPath = "c:\GitHub\office-js-snippets\samples"
$yamlFiles = Get-ChildItem -Path $samplesPath -Filter "*.yaml" -Recurse

# Initialize tracking structures
$issues = @{
    "MissingSpaceAfterComma" = @()
    "MissingSpaceInTemplate" = @()
    "MissingSpaceAfterColon" = @()
    "ArrayBracketSpacing" = @()
    "TrailingSpaces" = @()
}

$fileCount = 0
$totalIssues = 0

Write-Host "Analyzing $($yamlFiles.Count) YAML files..." -ForegroundColor Cyan
Write-Host ""

foreach ($file in $yamlFiles) {
    $fileCount++
    if ($fileCount % 50 -eq 0) {
        Write-Host "Processed $fileCount files..." -ForegroundColor Yellow
    }
    
    $content = Get-Content $file.FullName -Raw
    
    # Check if file has a script section
    if ($content -match 'script:\s+content:\s+\|-?\s+([\s\S]+?)(?=\n\w+:|$)') {
        $scriptContent = $matches[1]
        $lines = $scriptContent -split "`n"
        
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i]
            $lineNum = $i + 1
            
            # Pattern 1: Missing space after comma in arrays and function arguments
            # Look for patterns like: ,word or ," or ,\d but not ,\s or inside strings
            if ($line -match ',\S' -and $line -notmatch '^\s*//' -and $line -notmatch '//.*,\S') {
                # Simple check: comma followed by non-whitespace
                $issues["MissingSpaceAfterComma"] += [PSCustomObject]@{
                    File = $file.FullName.Replace($samplesPath, "samples")
                    Line = $lineNum
                    Code = $line.Trim()
                }
                $totalIssues++
            }
            
            # Pattern 2: Missing space after colon in template literals (e.g., key:${value})
            if ($line -match ':\$\{[^}]+\}' -and $line -notmatch ': \$\{') {
                $issues["MissingSpaceInTemplate"] += [PSCustomObject]@{
                    File = $file.FullName.Replace($samplesPath, "samples")
                    Line = $lineNum
                    Code = $line.Trim()
                }
                $totalIssues++
            }
            
            # Pattern 3: Trailing spaces (spaces at end of line, excluding empty lines)
            if ($line -match '\S\s+$') {
                $issues["TrailingSpaces"] += [PSCustomObject]@{
                    File = $file.FullName.Replace($samplesPath, "samples")
                    Line = $lineNum
                    Code = $line -replace '\s+$', '[SPACES]'
                }
                $totalIssues++
            }
            
            # Pattern 4: Array bracket spacing issues [item or item]
            if ($line -match '\[\S' -and $line -notmatch '\[\[' -and $line -notmatch '^\s*//' -and $line -notmatch '\["' -and $line -notmatch '//.*\[\S') {
                $issues["ArrayBracketSpacing"] += [PSCustomObject]@{
                    File = $file.FullName.Replace($samplesPath, "samples")
                    Line = $lineNum
                    Code = $line.Trim()
                }
                $totalIssues++
            }
            
            if ($line -match '\S\]' -and $line -notmatch '\]\]' -and $line -notmatch '^\s*//' -and $line -notmatch '"\]' -and $line -notmatch '//.*\S\]' -and $line -notmatch '\d\]') {
                $issues["ArrayBracketSpacing"] += [PSCustomObject]@{
                    File = $file.FullName.Replace($samplesPath, "samples")
                    Line = $lineNum
                    Code = $line.Trim()
                }
                $totalIssues++
            }
            
            # Pattern 5: Missing space after colon in object properties (key:value vs key: value)
            # But exclude URL patterns, case statements, ternary operators, and type annotations
            if ($line -match '\w:\w' -and 
                $line -notmatch '^\s*//' -and 
                $line -notmatch 'http[s]?:' -and 
                $line -notmatch '//http' -and
                $line -notmatch 'case\s+.*:' -and
                $line -notmatch '\?\s*\w+:\w+' -and
                $line -notmatch ':\s*Word\.' -and
                $line -notmatch ':\s*Excel\.' -and
                $line -notmatch ':\s*Office\.' -and
                $line -notmatch ':\s*string' -and
                $line -notmatch ':\s*number' -and
                $line -notmatch ':\s*boolean' -and
                $line -notmatch ':\s*void' -and
                $line -notmatch 'for\s*\(' -and
                $line -notmatch '\w+:\s*\w+\(' -and
                $line -notmatch '\w+:\d+') {
                
                # Additional check: not a type annotation pattern
                if ($line -notmatch '\w+:\s*\w+\.\w+' -and $line -notmatch ':\s*any' -and $line -notmatch ':\s*{') {
                    $issues["MissingSpaceAfterColon"] += [PSCustomObject]@{
                        File = $file.FullName.Replace($samplesPath, "samples")
                        Line = $lineNum
                        Code = $line.Trim()
                    }
                    $totalIssues++
                }
            }
        }
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "Analysis Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Total files analyzed: $fileCount" -ForegroundColor White
Write-Host "Total issues found: $totalIssues" -ForegroundColor Yellow
Write-Host ""

# Generate report
$reportPath = "c:\GitHub\office-js-snippets\style-issues-report.txt"
$report = @()

$report += "="*80
$report += "TypeScript Code Style Issues Report"
$report += "Generated: $(Get-Date)"
$report += "Files Analyzed: $fileCount"
$report += "Total Issues Found: $totalIssues"
$report += "="*80
$report += ""

foreach ($category in $issues.Keys | Sort-Object) {
    $categoryIssues = $issues[$category]
    if ($categoryIssues.Count -gt 0) {
        $report += ""
        $report += "-"*80
        $report += "ISSUE CATEGORY: $category"
        $report += "Total occurrences: $($categoryIssues.Count)"
        $report += "-"*80
        $report += ""
        
        # Group by file
        $byFile = $categoryIssues | Group-Object -Property File
        $report += "Files affected: $($byFile.Count)"
        $report += ""
        
        # Show first 20 files with issues
        $filesToShow = if ($byFile.Count -gt 20) { 20 } else { $byFile.Count }
        $report += "Showing details for first $filesToShow files:"
        $report += ""
        
        foreach ($fileGroup in ($byFile | Select-Object -First 20)) {
            $report += "  FILE: $($fileGroup.Name)"
            foreach ($issue in ($fileGroup.Group | Select-Object -First 5)) {
                $report += "    Line $($issue.Line): $($issue.Code)"
            }
            if ($fileGroup.Count -gt 5) {
                $report += "    ... and $($fileGroup.Count - 5) more occurrence(s) in this file"
            }
            $report += ""
        }
        
        if ($byFile.Count -gt 20) {
            $report += "  ... and $($byFile.Count - 20) more file(s) with this issue"
            $report += ""
        }
    }
}

$report += ""
$report += "="*80
$report += "Summary by Category"
$report += "="*80
foreach ($category in $issues.Keys | Sort-Object) {
    $report += "$($category): $($issues[$category].Count) occurrences in $(($issues[$category] | Select-Object -Unique File).Count) files"
}

$report | Out-File -FilePath $reportPath -Encoding UTF8
Write-Host "Full report saved to: $reportPath" -ForegroundColor Green
Write-Host ""

# Display summary on console
Write-Host "Summary by Category:" -ForegroundColor Cyan
Write-Host "-"*40
foreach ($category in $issues.Keys | Sort-Object) {
    $fileCount = ($issues[$category] | Select-Object -Unique File).Count
    Write-Host "$category`: $($issues[$category].Count) occurrences in $fileCount files" -ForegroundColor White
}
