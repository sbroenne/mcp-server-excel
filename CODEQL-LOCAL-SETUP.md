# CodeQL Local Setup and Execution Guide

## Installation

### Option 1: Direct Download (Recommended)
1. Download CodeQL CLI:
   https://github.com/github/codeql-cli-binaries/releases/latest
   
2. Extract to: C:\Tools\codeql
   
3. Add to PATH:
   $env:PATH += ';C:\Tools\codeql'
   [Environment]::SetEnvironmentVariable('PATH', $env:PATH, 'User')

### Option 2: GitHub CLI Extension
gh extension install github/gh-codeql

## Setup CodeQL Database

# Create database from C# code
codeql database create ./codeql-db --language=csharp --source-root=.

# Or use build command for compiled languages
codeql database create ./codeql-db --language=csharp --command='dotnet build -c Release'

## Run Analysis

# Download query packs
codeql pack download codeql/csharp-queries

# Run analysis with your config
codeql database analyze ./codeql-db 
  --format=sarif-latest 
  --output=./results.sarif 
  --sarif-category=csharp 
  codeql/csharp-queries:codeql-suites/csharp-security-and-quality.qls 
  -- --additional-packs=.github/codeql

## View Results

# Convert to human-readable format
codeql database interpret-results ./codeql-db ./results.sarif --format=csv --output=./results.csv

# Or view in VS Code with SARIF Viewer extension
code --install-extension MS-SarifVSCode.sarif-viewer
code ./results.sarif

## Quick Test (After Installation)

# Test your config without full scan
codeql resolve queries .github/codeql/codeql-config.yml

## Expected Time
- Database creation: 2-5 minutes
- Analysis: 3-10 minutes
- Total: ~5-15 minutes

## Alternative: Use GitHub Actions Locally

# Install act (runs GitHub Actions locally)
winget install nektos.act

# Run CodeQL workflow locally
act pull_request --workflows .github/workflows/codeql.yml
