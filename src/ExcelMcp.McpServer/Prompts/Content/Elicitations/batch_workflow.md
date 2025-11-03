# BEFORE STARTING BATCH WORKFLOW - GATHER THIS INFO

REQUIRED:
☐ Excel file path (which file to batch process)
☐ List of operations to perform (what needs to be done)

BATCH MODE DETECTION KEYWORDS:
☐ Numbers: "import 4 queries", "create 5 parameters", "3 measures"
☐ Plurals: "queries", "parameters", "measures", "relationships", "worksheets"
☐ Lists: User provides enumerated items (Sales, Products, Customers)
☐ Repetitive: "each", "all", "every"

WORKFLOW PLANNING:
☐ How many operations total? (2+ = use batch mode)
☐ What's the sequence? (plan order: create → configure → populate)
☐ Save at end or discard? (commit_excel_batch save: true/false)

PREREQUISITES:
☐ File already exists? (use excel_file create-empty if not)
☐ Dependencies resolved? (e.g., data loaded before creating measures)

EXPECTED PATTERN:
1. begin_excel_batch(excelPath) → Save batchId
2. Operation 1 with batchId
3. Operation 2 with batchId
4. ... (repeat for all operations)
5. commit_excel_batch(batchId, save: true)

PERFORMANCE BENEFITS:
☐ 2 operations: 50-60% faster
☐ 4 operations: 75-85% faster
☐ 10+ operations: 90-95% faster

ASK USER to confirm all operations before starting batch workflow.
REMIND USER to use batchId for EVERY operation in the batch.
