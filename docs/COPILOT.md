# GitHub Copilot Integration Guide

Complete guide for using ExcelMcp with GitHub Copilot for Excel automation.

## Configure Your IDE for Optimal ExcelMcp Development

### VS Code Settings

Add to your `settings.json`:

```json
{
  "github.copilot.enable": {
    "*": true,
    "csharp": true,
    "markdown": true
  },
  "github.copilot.advanced": {
    "listCount": 10,
    "inlineSuggestCount": 3
  }
}
```

## Enable ExcelMcp Support in Your Projects

To make GitHub Copilot aware of ExcelMcp in your own projects:

1. **Copy the Copilot Instructions** to your project:

   ```bash
   # Copy ExcelMcp automation instructions to your project's .github directory
   curl -o .github/excel-powerquery-vba-instructions.md https://raw.githubusercontent.com/sbroenne/mcp-server-excel/main/.github/excel-powerquery-vba-instructions.md
   ```

2. **Configure VS Code** (optional but recommended):

   ```json
   {
     "github.copilot.enable": {
       "*": true,
       "csharp": true,
       "powershell": true,
       "yaml": true
     }
   }
   ```

## Effective Copilot Prompting

With the ExcelMcp instructions installed, Copilot will automatically suggest ExcelMcp commands. Here's how to get the best results:

### General Prompting Tips

```text
"Use ExcelMcp to..." - Start prompts this way for targeted suggestions
"Create a complete workflow using ExcelMcp that..." - For end-to-end automation
"Help me troubleshoot this ExcelMcp command..." - For debugging assistance
```

### Reference the Instructions

The ExcelMcp instruction file (`.github/excel-powerquery-vba-instructions.md`) contains complete workflow examples for:

- Data Pipeline automation
- VBA automation workflows  
- Combined PowerQuery + VBA scenarios
- Report generation patterns

Copilot will reference these automatically when you mention ExcelMcp in your prompts.

## Essential Copilot Prompts for ExcelMcp

### Extract Power Query M Code from Excel

```text
Use ExcelMcp pq-list to show all Power Queries embedded in my Excel workbook
Extract M code with pq-export so Copilot can analyze my data transformations
Use pq-view to display the hidden M formula code from my Excel file
Check what data sources are available with pq-sources command
```

### Debug & Validate Power Query

```text
Use ExcelMcp pq-errors to check for issues in my Excel Power Query
Validate M code syntax with pq-verify before updating my Excel file
Test Power Query data preview with pq-peek to see sample results
Use pq-test to verify my query connections work properly
```

### Advanced Excel Automation

```text
Use ExcelMcp to refresh pq-refresh then sheet-read to extract updated data
Load connection-only queries to worksheets with pq-loadto command
Manage cell formulas with cell-get-formula and cell-set-formula commands
Run VBA macros with script-run and check results with sheet-read commands
Export VBA scripts with script-export for complete Excel code backup
Use setup-vba-trust to configure VBA access for automated workflows
Create macro-enabled workbooks with create-empty "file.xlsm" for VBA support
```

## Advanced Copilot Techniques

### Context-Aware Code Generation

When working with ExcelCLI, provide context to Copilot:

```text
I'm working with ExcelMcp to process Excel files. 
I need to:
- Read data from multiple worksheets
- Combine data using Power Query
- Apply business logic with VBA
- Export results to CSV

Generate a complete PowerShell script using ExcelMcp commands.
```

### Error Handling Patterns

Ask Copilot to generate robust error handling:

```text
Create error handling for ExcelMcp commands that:
- Checks if Excel files exist
- Validates VBA trust configuration
- Handles Excel COM errors gracefully
- Provides meaningful error messages
```

### Performance Optimization

Ask Copilot for performance improvements:

```text
Optimize this ExcelMcp workflow for processing large datasets:
- Minimize Excel file operations
- Use efficient Power Query patterns
- Implement parallel processing where possible
```

## Troubleshooting with Copilot

### Common Issues

Ask Copilot to help diagnose:

```text
ExcelMcp pq-refresh is failing with "connection error"
Help me debug this Power Query issue and suggest fixes
```

```text
VBA script-run command returns "access denied"
Help me troubleshoot VBA trust configuration issues
```

```text
Excel processes are not cleaning up after ExcelMcp commands
Help me identify and fix process cleanup issues
```

### Best Practices

Copilot can suggest best practices:

```text
What are the best practices for using ExcelMcp in CI/CD pipelines?
How should I structure ExcelMcp commands for maintainable automation scripts?
What error handling patterns should I use with ExcelCLI?
```

## Integration with Other Tools

### PowerShell Modules

Ask Copilot to create PowerShell wrappers:

```text
Create a PowerShell module that wraps ExcelMcp commands with:
- Parameter validation
- Error handling
- Logging
- Progress reporting
```

### Azure DevOps Integration

```text
Create Azure DevOps pipeline tasks that use ExcelMcp to:
- Process Excel reports in build pipelines
- Generate data exports for deployment
- Validate Excel file formats and content
```

This guide enables developers to leverage GitHub Copilot's full potential when working with ExcelMcp for Excel automation, making the development process more efficient and the resulting code more robust.
