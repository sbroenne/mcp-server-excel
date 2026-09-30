"""Identical user requests for the CLI and MCP saved-workbook evaluations."""

from __future__ import annotations


def chart_prompt(path: str, *, below: bool) -> str:
    if below:
        return f"""
Create a workbook at {path}. On Sheet1, put this data in A1:C6:
Month, Revenue, Expenses
January, 50000, 35000
February, 55000, 38000
March, 48000, 32000
April, 62000, 41000
May, 58000, 39000

Create a clustered column chart with one series each for Revenue and Expenses
and the month names as category labels. Place it below row 6 without overlapping
the source data. Save and close the workbook, and report its chart position.
"""
    return f"""
Create a workbook at {path}. On Sheet1, put this data in A1:D5:
Product, Q1, Q2, Q3
Widget, 100, 150, 120
Gadget, 80, 90, 110
Device, 200, 180, 220
Tool, 50, 60, 75

Convert the data to an Excel Table named ProductSales. Create a line chart with
one series per quarter and product names as category labels. Place it to the
right of the Table without overlapping the data. Save and close the workbook.
"""


def pivot_slicer_prompt(path: str) -> str:
    return f"""
Create a workbook at {path}. On Sheet1 enter:
Region, Product, Quarter, Sales
North, Laptop, Q1, 15000
North, Phone, Q1, 8000
North, Laptop, Q2, 18000
North, Phone, Q2, 9500
South, Laptop, Q1, 12000
South, Phone, Q1, 7500
South, Laptop, Q2, 14000
South, Phone, Q2, 8200

Convert the data to a Table named SalesData. Create a PivotTable on a new
Analysis sheet with Region as rows and total Sales as values.
Create a Region slicer on Analysis at E2 and filter it to North.
Create a temporary Product slicer at G2, then remove that temporary slicer.
Keep the Region slicer and its North selection. Save and close the workbook.
Report the filtered sales total.
"""


def table_slicer_prompt(path: str) -> str:
    return f"""
Create a workbook at {path}. On Sheet1 enter:
Department, Employee, Status, Salary
Engineering, Alice, Active, 85000
Engineering, Bob, Active, 92000
Marketing, Carol, Active, 78000
Marketing, Dave, Inactive, 70000
Sales, Eve, Active, 65000
Sales, Frank, Inactive, 62000
Engineering, Grace, Active, 88000
Sales, Henry, Active, 71000

Convert the data to an Excel Table named Employees. Create a Department Table
slicer at F2 and select Engineering. Create a Status Table slicer at H2 and
select Active. Keep both slicers and their filters. Save and close the workbook.
Report the visible employees and their total salary.
"""


def combined_slicer_prompt(path: str) -> str:
    return f"""
Create a workbook at {path}. On Sheet1 enter:
Category, Product, Warehouse, Stock, Price
Electronics, Laptop, West, 50, 999
Electronics, Phone, West, 120, 599
Electronics, Laptop, East, 35, 999
Electronics, Phone, East, 80, 599
Furniture, Desk, West, 25, 350
Furniture, Chair, West, 40, 175
Furniture, Desk, East, 30, 350
Furniture, Chair, East, 55, 175

Convert the data to a Table named Inventory. Create a PivotTable on a new
Summary sheet with Category as rows and total Stock as values.
Create a Warehouse Table slicer on Sheet1 at G2 and select West.
Create a Category PivotTable slicer on Summary at D2 and select Electronics.
Keep both slicers and their filters. Save and close the workbook.
Report the visible Table rows separately from the PivotTable total; explain
which data each slicer filters.
"""
