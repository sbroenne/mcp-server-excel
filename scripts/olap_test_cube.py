"""Synthetic external OLAP cube for ExcelMcp integration tests.

Serves a small Atoti cube over XMLA on localhost. Contains no customer data.
Started by Start-OlapTestCube.ps1. Usage: python olap_test_cube.py <port>
"""

import sys
import time

import atoti as tt
import pandas as pd

port = int(sys.argv[1])
months = [f"{year}-{month:02d}" for year in (2024, 2025) for month in range(1, 13)]
sales = pd.DataFrame(
    {
        "Id": range(len(months)),
        "Year": [month[:4] for month in months],
        "Month": months,
        "Amount": [float(index + 1) for index in range(len(months))],
    }
)

session = tt.Session.start(tt.SessionConfig(port=port))
table = session.read_pandas(sales, table_name="Sales", keys={"Id"})
cube = session.create_cube(table, "SalesCube")
hierarchies, levels = cube.hierarchies, cube.levels
hierarchies["Calendar"] = [levels["Year"], levels["Month"]]
del hierarchies["Year"], hierarchies["Month"]

print("READY", flush=True)
while True:
    time.sleep(60)
