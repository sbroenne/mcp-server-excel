# BEFORE CREATING DAX MEASURES - GATHER THIS INFO

REQUIRED FOR EACH MEASURE:
☐ Measure name (e.g., 'Total Sales', 'Avg Price', 'Customer Count')
☐ Target table (which Data Model table owns this measure)
☐ DAX formula (e.g., 'SUM(Sales[Amount])', 'AVERAGE(Sales[Price])')

RECOMMENDED:
☐ Format string:
  - '#,##0.00' for decimals
  - '$#,##0' for currency
  - '0.00%' for percentage
  - 'General Number' for general
☐ Display folder (organize measures in categories like 'Revenue', 'Orders')
☐ Description (helps other users understand purpose)

WORKFLOW OPTIMIZATION:
☐ Are you creating 2+ measures? → Use batch mode (begin_excel_batch)
☐ Do Data Model tables exist? → Check with excel_datamodel(action: 'list-tables') first
☐ Is data loaded to Data Model? → Queries must use loadDestination: 'data-model' or 'both'

ASK USER FOR MISSING INFO.
BATCH MODE saves 75-95% time for multiple measures.
