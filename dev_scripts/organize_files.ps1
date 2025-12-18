# Script to organize test and debug files

# Move unit test files
$unitTests = @(
    "test_search.py", "test_search_fixed.py", "test_search_terms.py",
    "test_search_formats.py", "test_corrected_criteria.py", "test_final_search.py",
    "test_direct_search.py", "test_early_termination_fix.py"
)

foreach ($file in $unitTests) {
    if (Test-Path $file) {
        Move-Item -Path $file -Destination "tests\unit\"
        Write-Host "Moved $file to tests\unit\"
    }
}

# Move integration test files
$integrationTests = @(
    "test_performance_comparison.py", "test_optimized_list.py",
    "test_dynamic_limit.py", "test_dynamic_limit_fixed.py", "test_search_days.py"
)

foreach ($file in $integrationTests) {
    if (Test-Path $file) {
        Move-Item -Path $file -Destination "tests\integration\"
        Write-Host "Moved $file to tests\integration\"
    }
}

# Move debug/analysis scripts
$debugScripts = @(
    "debug_dates.py", "debug_date_filtering.py", "debug_dates_full.py",
    "debug_search_criteria.py", "check_emails.py", "check_approval_emails.py",
    "check_total_emails.py", "check_1000th_email.py", "analyze_distribution.py",
    "analyze_email_distribution.py", "find_7day_cutoff.py"
)

foreach ($file in $debugScripts) {
    if (Test-Path $file) {
        Move-Item -Path $file -Destination "tests\scripts\"
        Write-Host "Moved $file to tests\scripts\"
    }
}

Write-Host "File organization complete!"