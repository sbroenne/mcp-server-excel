"""Extract a small, attributed WDI snapshot. Never reads or writes workbook files."""

import argparse
import csv
import hashlib
import io
import json
import math
from pathlib import Path
import zipfile

COUNTRIES = (
    "ARG AUS BRA CAN CHL CHN DEU EGY ESP FRA GBR IDN IND ITA JPN KOR "
    "MEX NGA NOR NZL POL SWE TUR USA ZAF"
).split()
INDICATORS = {
    "NY.GDP.PCAP.PP.KD": ("Income", "Constant 2021 international $ per person"),
    "NY.GDP.MKTP.KD": ("GDP", "Constant 2015 US$"),
    "NY.GDP.MKTP.KD.ZG": ("Growth", "Annual percentage points"),
    "FP.CPI.TOTL.ZG": ("Inflation", "Annual percentage points"),
    "SP.POP.TOTL": ("Population", "People"),
    "SP.DYN.LE00.IN": ("Life", "Years"),
    "EG.ELC.ACCS.ZS": ("Electricity", "Percent of population"),
    "IT.NET.USER.ZS": ("Internet", "Percent of population"),
    "NE.TRD.GNFS.ZS": ("Trade", "Percent of GDP"),
}
SOURCE = "https://databankfiles.worldbank.org/public/ddpext_download/WDI_CSV.zip"


def csv_rows(archive, name):
    return csv.DictReader(io.TextIOWrapper(archive.open(name), encoding="utf-8-sig"))


def write_csv(path, fields, rows):
    with path.open("w", newline="", encoding="utf-8") as output:
        writer = csv.DictWriter(output, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)


def prepare(archive_path, output, retrieved):
    output.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(archive_path) as archive:
        metadata = {}
        for row in csv_rows(archive, "WDISeries.csv"):
            code = row["Series Code"]
            if code not in INDICATORS:
                continue
            if row["License Type"] != "CC BY-4.0":
                raise ValueError(f"Indicator {code} is not licensed CC BY-4.0")
            metadata[code] = {
                "Code": code, "Metric": INDICATORS[code][0],
                "Indicator": row["Indicator Name"], "Unit": INDICATORS[code][1],
                "License": row["License Type"], "Source": row["Source"],
                "Definition": row["Long definition"],
                "Limitations": row["Limitations and exceptions"],
                "Aggregation": row["Aggregation method"],
            }
        if set(metadata) != set(INDICATORS):
            raise ValueError("Missing requested indicator metadata")

        countries = []
        for row in csv_rows(archive, "WDICountry.csv"):
            if row["Country Code"] in COUNTRIES:
                if not row["Region"]:
                    raise ValueError("Aggregates must not enter country comparisons")
                countries.append({
                    "Code": row["Country Code"], "Country": row["Short Name"],
                    "Region": row["Region"], "IncomeGroup": row["Income Group"],
                })
        if {row["Code"] for row in countries} != set(COUNTRIES):
            raise ValueError("Missing requested countries")
        countries.sort(key=lambda row: row["Country"])
        observations = []
        seen = set()
        for row in csv_rows(archive, "WDICSV.csv"):
            country, code = row["Country Code"], row["Indicator Code"]
            if country not in COUNTRIES or code not in INDICATORS:
                continue
            if (country, code) in seen:
                raise ValueError(f"Duplicate series: {country} {code}")
            seen.add((country, code))
            for year in range(2000, 2025):
                text = row[str(year)].strip()
                value = float(text) if text else None
                if value is not None and not math.isfinite(value):
                    raise ValueError("Non-finite source observation")
                observations.append({
                    "CountryCode": country, "Year": year,
                    "IndicatorCode": code, "Metric": INDICATORS[code][0],
                    "Value": value,
                })
        if len(seen) != len(COUNTRIES) * len(INDICATORS):
            raise ValueError("Missing country/indicator series")
        observations.sort(key=lambda row: (row["CountryCode"], row["Year"], row["Metric"]))
        write_csv(output / "observations.csv", list(observations[0]), observations)
        write_csv(output / "countries.csv", list(countries[0]), countries)
        write_csv(output / "indicators.csv", list(next(iter(metadata.values()))), metadata.values())

    with archive_path.open("rb") as source:
        digest = hashlib.file_digest(source, "sha256").hexdigest()
    coverage = {
        metric: {
            str(year): sum(
                row["Value"] is not None for row in observations
                if row["Year"] == year and row["Metric"] == metric
            ) for year in (2000, 2010, 2020, 2023, 2024)
        } for metric, _ in INDICATORS.values()
    }
    manifest = {
        "dataset": "World Development Indicators",
        "publisher": "World Bank", "source": SOURCE,
        "catalogue": "https://datacatalog.worldbank.org/search/dataset/0037712/world-development-indicators",
        "terms": "https://datacatalog.worldbank.org/public-licenses#cc-by",
        "license": "https://creativecommons.org/licenses/by/4.0/",
        "attribution": "Source: World Bank, World Development Indicators (CC BY 4.0). Selected and reshaped for this sample; no World Bank endorsement.",
        "retrieved": retrieved, "archiveSha256": digest,
        "years": [2000, 2024], "countries": len(countries),
        "indicators": len(metadata), "observations": len(observations),
        "missing": sum(row["Value"] is None for row in observations),
        "coverage": coverage,
        "notes": [
            "Country region and income-group labels reflect the downloaded metadata, not historical classifications.",
            "Rates are percentage-point values as supplied, not decimal fractions.",
            "Missing values remain missing. No interpolation or zero replacement.",
            "The sample countries are not the whole world. No country average is labelled as a World Bank regional aggregate.",
        ],
    }
    (output / "sources.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(manifest, indent=2))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("archive", type=Path)
    parser.add_argument("--output", type=Path, default=Path(__file__).parent / "data")
    parser.add_argument("--retrieved", required=True, help="Actual retrieval date in ISO format")
    args = parser.parse_args()
    prepare(args.archive, args.output, args.retrieved)
