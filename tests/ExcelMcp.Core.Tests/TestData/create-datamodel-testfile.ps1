# Script to create a test Excel file with Data Model for TOM tests
# This creates a real Data Model workbook that can be committed to git

$testFile = "D:\source\mcp-server-excel\tests\ExcelMcp.Core.Tests\TestData\DataModelSample.xlsx"
$tempDir = "D:\source\mcp-server-excel\tests\ExcelMcp.Core.Tests\TestData\temp"

# Clean up if exists
if (Test-Path $testFile) { Remove-Item $testFile -Force }
if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

Write-Host "Step 1: Creating empty workbook..." -ForegroundColor Cyan
dotnet run --project src/ExcelMcp.CLI/ExcelMcp.CLI.csproj -- create-empty $testFile

Write-Host "`nStep 2: Creating worksheets..." -ForegroundColor Cyan
dotnet run --project src/ExcelMcp.CLI/ExcelMcp.CLI.csproj -- sheet-create $testFile Sales
dotnet run --project src/ExcelMcp.CLI/ExcelMcp.CLI.csproj -- sheet-create $testFile Customers
dotnet run --project src/ExcelMcp.CLI/ExcelMcp.CLI.csproj -- sheet-create $testFile Products

Write-Host "`nStep 3: Adding Sales data..." -ForegroundColor Cyan
@"
SalesID,Date,CustomerID,ProductID,Amount,Quantity
1,2024-01-15,101,1001,1500.00,3
2,2024-01-20,102,1002,2200.00,2
3,2024-02-10,103,1003,750.00,1
4,2024-02-15,101,1001,3000.00,6
5,2024-03-05,104,1004,1200.00,2
6,2024-03-10,102,1003,900.00,1
7,2024-04-01,105,1002,2800.00,4
8,2024-04-15,103,1004,1600.00,2
9,2024-05-05,104,1001,2100.00,5
10,2024-05-20,105,1003,950.00,1
"@ | Out-File "$tempDir\sales.csv" -Encoding UTF8

dotnet run --project src/ExcelMcp.CLI/ExcelMcp.CLI.csproj -- sheet-write $testFile Sales "$tempDir\sales.csv"

Write-Host "`nStep 4: Adding Customers data..." -ForegroundColor Cyan
@"
CustomerID,CustomerName,Segment,Country
101,Acme Corp,Enterprise,USA
102,Beta Industries,SMB,Canada
103,Gamma Solutions,Enterprise,USA
104,Delta Services,SMB,UK
105,Epsilon Group,Enterprise,Germany
"@ | Out-File "$tempDir\customers.csv" -Encoding UTF8

dotnet run --project src/ExcelMcp.CLI/ExcelMcp.CLI.csproj -- sheet-write $testFile Customers "$tempDir\customers.csv"

Write-Host "`nStep 5: Adding Products data..." -ForegroundColor Cyan
@"
ProductID,ProductName,Category,Price
1001,Widget Pro,Hardware,500.00
1002,Software Suite,Software,1100.00
1003,Service Plan,Services,750.00
1004,Premium Package,Bundle,800.00
"@ | Out-File "$tempDir\products.csv" -Encoding UTF8

dotnet run --project src/ExcelMcp.CLI/ExcelMcp.CLI.csproj -- sheet-write $testFile Products "$tempDir\products.csv"

Write-Host "`nStep 5: Creating PowerQuery to load Sales to Data Model..." -ForegroundColor Cyan
@"
let
    Source = Excel.CurrentWorkbook(){[Name="Sales"]}[Content],
    ChangedType = Table.TransformColumnTypes(Source,{
        {"SalesID", Int64.Type},
        {"Date", type datetime},
        {"CustomerID", Int64.Type},
        {"ProductID", Int64.Type},
        {"Amount", type number},
        {"Quantity", Int64.Type}
    })
in
    ChangedType
"@ | Out-File "$tempDir\sales.pq" -Encoding UTF8

dotnet run --project src/ExcelMcp.CLI/ExcelMcp.CLI.csproj -- pq-import $testFile SalesData "$tempDir\sales.pq" --privacy-level Private

Write-Host "`nStep 6: Creating PowerQuery to load Customers to Data Model..." -ForegroundColor Cyan
@"
let
    Source = Excel.CurrentWorkbook(){[Name="Customers"]}[Content],
    ChangedType = Table.TransformColumnTypes(Source,{
        {"CustomerID", Int64.Type},
        {"CustomerName", type text},
        {"Segment", type text},
        {"Country", type text}
    })
in
    ChangedType
"@ | Out-File "$tempDir\customers.pq" -Encoding UTF8

dotnet run --project src/ExcelMcp.CLI/ExcelMcp.CLI.csproj -- pq-import $testFile CustomersData "$tempDir\customers.pq" --privacy-level Private

Write-Host "`nStep 7: Creating PowerQuery to load Products to Data Model..." -ForegroundColor Cyan
@"
let
    Source = Excel.CurrentWorkbook(){[Name="Products"]}[Content],
    ChangedType = Table.TransformColumnTypes(Source,{
        {"ProductID", Int64.Type},
        {"ProductName", type text},
        {"Category", type text},
        {"Price", type number}
    })
in
    ChangedType
"@ | Out-File "$tempDir\products.pq" -Encoding UTF8

dotnet run --project src/ExcelMcp.CLI/ExcelMcp.CLI.csproj -- pq-import $testFile ProductsData "$tempDir\products.pq" --privacy-level Private

Write-Host "`nStep 8: Setting queries to load to Data Model..." -ForegroundColor Cyan
# Note: This step may require manual intervention in Excel
# Open the file and:
# 1. Data > Queries & Connections
# 2. Right-click each query (SalesData, CustomersData, ProductsData)
# 3. Select "Load To..." > "Only Create Connection" > Check "Add this data to the Data Model"

Write-Host "`n✓ Base file created: $testFile" -ForegroundColor Green
Write-Host "`nMANUAL STEPS REQUIRED:" -ForegroundColor Yellow
Write-Host "1. Open $testFile in Excel" -ForegroundColor Yellow
Write-Host "2. Data > Queries & Connections" -ForegroundColor Yellow
Write-Host "3. For each query (SalesData, CustomersData, ProductsData):" -ForegroundColor Yellow
Write-Host "   - Right-click > Load To..." -ForegroundColor Yellow
Write-Host "   - Select 'Only Create Connection'" -ForegroundColor Yellow
Write-Host "   - Check 'Add this data to the Data Model'" -ForegroundColor Yellow
Write-Host "4. Save and close" -ForegroundColor Yellow

# Clean up temp files
Remove-Item $tempDir -Recurse -Force

Write-Host "`nReady to open file for manual configuration? (Press Enter)" -ForegroundColor Cyan
Read-Host
Start-Process $testFile
