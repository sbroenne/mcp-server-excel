import json, glob, os

target_tests = [
    "test_cli_file_and_worksheet_workflow",
    "test_cli_financial_report_automation",
    "test_cli_star_schema_workflow",
    "test_cli_powerquery_products_workflow",
    "test_cli_sales_report_workflow",
    "test_cli_pivottable_slicer_workflow",
    "test_cli_table_slicer_workflow",
]

# Find the most recent file with CLI tests
files = sorted(glob.glob("aitest-reports/results_pytest-aitest_*.json"), reverse=True)

for f in files:
    with open(f) as fh:
        data = json.load(fh)
    tests = data.get("tests", data.get("results", []))
    cli_tests = [t for t in tests if "cli" in t.get("name", "").lower()]
    if len(cli_tests) >= 7:
        print(f"=== REPORT FILE: {os.path.basename(f)} ===")
        print(f"Total tests: {len(tests)}, CLI tests: {len(cli_tests)}\n")
        
        for target in target_tests:
            matches = [t for t in tests if target in t.get("name", "")]
            if matches:
                t = matches[0]
                print(f"{'='*80}")
                print(f"TEST: {t['name']}")
                print(f"STATUS: {t['status']}")
                print(f"FAILURE REASON: {t.get('failure_reason', 'N/A')}")
                
                # Print tool calls if available
                tool_calls = t.get("tool_calls", [])
                if tool_calls:
                    print(f"\nTOOL CALLS ({len(tool_calls)}):")
                    for i, tc in enumerate(tool_calls):
                        if isinstance(tc, dict):
                            # Show exit codes and errors
                            output = tc.get("output", tc.get("result", ""))
                            if isinstance(output, str) and ("exit_code" in output or "error" in output or "Error" in output):
                                print(f"  Call {i}: {json.dumps(tc, indent=2)[:500]}")
                
                # Print conversation excerpts if available
                conversation = t.get("conversation", t.get("messages", []))
                if conversation:
                    print(f"\nCONVERSATION EXCERPTS (last 3):")
                    for msg in conversation[-3:]:
                        role = msg.get("role", "?")
                        content = str(msg.get("content", ""))[:300]
                        print(f"  [{role}]: {content}")
                print()
            else:
                print(f"TEST: {target} - NOT FOUND in report")
        break
else:
    print("No report with 7+ CLI tests found")
