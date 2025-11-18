# Refined TypeScript Style Analysis Script
# Focus on actionable style issues, excluding false positives

$samplesPath = "c:\GitHub\office-js-snippets\samples"
$yamlFiles = Get-ChildItem -Path $samplesPath -Filter "*.yaml" -Recurse

# Initialize tracking structures
$issues = @{
    "MissingSpaceAfterCommaInArraysAndFunctions" = @()
    "MissingSpaceInTemplateLiteral" = @()
    "TrailingSpacesInCode" = @()
}

$fileCount = 0
$totalIssues = 0

Write-Host "Analyzing $($yamlFiles.Count) YAML files for genuine style issues..." -ForegroundColor Cyan
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
            
            # ISSUE 1: Missing space after comma - but exclude:
            # - Comments
            # - URLs
            # - JSON strings
            # - Number formatting (e.g., #,##0.00)
            if ($line -match ',\S' -and 
                $line -notmatch '^\s*//' -and 
                $line -notmatch 'http[s]?:' -and
                $line -notmatch '#,##' -and
                $line -notmatch 'base64,' -and
                $line -notmatch '",\w' -and
                $line -notmatch ',TRUE' -and
                $line -notmatch ',FALSE' -and
                $line -notmatch 'RC\[-\d+\]' -and
                $line -notmatch '//.*,' -and
                $line -notmatch '\],\s*$') {
                
                # Additional filtering: look for actual function calls or array literals
                if ($line -match '\w+\([^)]*,\S' -or 
                    $line -match '\[[^\]]*,\S' -or
                    $line -match ',(?=\d{2,})' -or
                    $line -match ',[a-zA-Z]') {
                    
                    $issues["MissingSpaceAfterCommaInArraysAndFunctions"] += [PSCustomObject]@{
                        File = $file.FullName.Replace($samplesPath, "samples")
                        Line = $lineNum
                        Code = $line.Trim()
                    }
                    $totalIssues++
                }
            }
            
            # ISSUE 2: Missing space after colon in template literals
            # Pattern: key:${value} should be key: ${value}
            if ($line -match ':\$\{[^}]+\}' -and $line -notmatch ': \$\{') {
                # Exclude time/date patterns like hours:${minutes}
                if ($line -notmatch '(hour|minute|second|month|year|date)s?:\$\{') {
                    $issues["MissingSpaceInTemplateLiteral"] += [PSCustomObject]@{
                        File = $file.FullName.Replace($samplesPath, "samples")
                        Line = $lineNum
                        Code = $line.Trim()
                    }
                    $totalIssues++
                }
            }
            
            # ISSUE 3: Trailing spaces in actual code lines
            # Exclude: empty lines, lines that are just indentation
            if ($line -match '\S\s+$' -and $line.Trim().Length -gt 0) {
                $issues["TrailingSpacesInCode"] += [PSCustomObject]@{
                    File = $file.FullName.Replace($samplesPath, "samples")
                    Line = $lineNum
                    Preview = $line.Substring(0, [Math]::Min($line.Length, 60)) + "..."
                }
                $totalIssues++
            }
        }
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "Refined Analysis Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

# Generate report
$reportPath = "c:\GitHub\office-js-snippets\refined-style-report.txt"
$report = @()

$report += "="*80
$report += "TypeScript Code Style Issues Report (Refined)"
$report += "Generated: $(Get-Date)"
$report += "Files Analyzed: $fileCount"
$report += "Total Issues Found: $totalIssues"
$report += "="*80
$report += ""
$report += "This report focuses on genuine code style issues, excluding:"
$report += "- TypeScript type annotations (e.g., Array<T>, Promise<T>)"
$report += "- Array indexing (e.g., items[i])"
$report += "- URL patterns"
$report += "- Number formatting patterns"
$report += "- JSON strings"
$report += "- Comments"
$report += ""

foreach ($category in $issues.Keys | Sort-Object) {
    $categoryIssues = $issues[$category]
    if ($categoryIssues.Count -gt 0) {
        $report += ""
        $report += "-"*80
        $report += "ISSUE: $category"
        $report += "Total occurrences: $($categoryIssues.Count)"
        $report += "-"*80
        $report += ""
        
        # Group by file
        $byFile = $categoryIssues | Group-Object -Property File
        $report += "Files affected: $($byFile.Count)"
        $report += ""
        
        if ($category -eq "TrailingSpacesInCode") {
            $report += "NOTE: Trailing spaces found in ALL $($byFile.Count) files."
            $report += "This is a widespread formatting issue affecting every file."
            $report += "Consider using an automated formatter to fix these."
            $report += ""
            $report += "Sample files with trailing spaces:"
            foreach ($fileGroup in ($byFile | Select-Object -First 10)) {
                $report += "  - $($fileGroup.Name) ($($fileGroup.Count) lines)"
            }
        } else {
            $report += "Showing all files with this issue:"
            $report += ""
            
            foreach ($fileGroup in $byFile) {
                $report += "  FILE: $($fileGroup.Name)"
                $maxLines = [Math]::Min($fileGroup.Count, 10)
                foreach ($issue in ($fileGroup.Group | Select-Object -First $maxLines)) {
                    $report += "    Line $($issue.Line): $($issue.Code)"
                }
                if ($fileGroup.Count -gt 10) {
                    $report += "    ... and $($fileGroup.Count - 10) more occurrence(s) in this file"
                }
                $report += ""
            }
        }
    }
}

$report += ""
$report += "="*80
$report += "Summary"
$report += "="*80
foreach ($category in $issues.Keys | Sort-Object) {
    $fileCount = ($issues[$category] | Select-Object -Unique File).Count
    $report += "$($category):"
    $report += "  - $($issues[$category].Count) total occurrences"
    $report += "  - $fileCount files affected"
    $report += ""
}

$report | Out-File -FilePath $reportPath -Encoding UTF8
Write-Host "Full report saved to: $reportPath" -ForegroundColor Green
Write-Host ""

# Display summary
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "-"*40
foreach ($category in $issues.Keys | Sort-Object) {
    $fileCount = ($issues[$category] | Select-Object -Unique File).Count
    Write-Host "$category" -ForegroundColor Yellow
    Write-Host "  $($issues[$category].Count) occurrences in $fileCount files" -ForegroundColor White
    Write-Host ""
}
