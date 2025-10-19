using Xunit;
using Sbroenne.ExcelMcp.Core.Models;
using System.Collections.Generic;

namespace Sbroenne.ExcelMcp.Core.Tests.Models;

/// <summary>
/// Unit tests for Result types - no Excel required
/// Tests verify proper construction and serialization of Result objects
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
public class ResultTypesTests
{
    [Fact]
    public void OperationResult_Success_HasCorrectProperties()
    {
        // Arrange & Act
        var result = new OperationResult
        {
            Success = true,
            FilePath = "test.xlsx",
            Action = "create",
            ErrorMessage = null
        };

        // Assert
        Assert.True(result.Success);
        Assert.Equal("test.xlsx", result.FilePath);
        Assert.Equal("create", result.Action);
        Assert.Null(result.ErrorMessage);
    }

    [Fact]
    public void OperationResult_Failure_HasErrorMessage()
    {
        // Arrange & Act
        var result = new OperationResult
        {
            Success = false,
            FilePath = "test.xlsx",
            Action = "delete",
            ErrorMessage = "File not found"
        };

        // Assert
        Assert.False(result.Success);
        Assert.Equal("File not found", result.ErrorMessage);
    }

    [Fact]
    public void CellValueResult_WithValue_HasCorrectProperties()
    {
        // Arrange & Act
        var result = new CellValueResult
        {
            Success = true,
            FilePath = "test.xlsx",
            CellAddress = "A1",
            Value = "Hello",
            Formula = null,
            ValueType = "String"
        };

        // Assert
        Assert.True(result.Success);
        Assert.Equal("A1", result.CellAddress);
        Assert.Equal("Hello", result.Value);
        Assert.Equal("String", result.ValueType);
    }

    [Fact]
    public void CellValueResult_WithFormula_HasFormulaAndValue()
    {
        // Arrange & Act
        var result = new CellValueResult
        {
            Success = true,
            FilePath = "test.xlsx",
            CellAddress = "B1",
            Value = "42",
            Formula = "=SUM(A1:A10)",
            ValueType = "Number"
        };

        // Assert
        Assert.Equal("=SUM(A1:A10)", result.Formula);
        Assert.Equal("42", result.Value);
    }

    [Fact]
    public void ParameterListResult_WithParameters_HasCorrectStructure()
    {
        // Arrange & Act
        var result = new ParameterListResult
        {
            Success = true,
            FilePath = "test.xlsx",
            Parameters = new List<ParameterInfo>
            {
                new() { Name = "StartDate", Value = "2024-01-01", RefersTo = "Settings!A1" },
                new() { Name = "EndDate", Value = "2024-12-31", RefersTo = "Settings!A2" }
            }
        };

        // Assert
        Assert.True(result.Success);
        Assert.Equal(2, result.Parameters.Count);
        Assert.Equal("StartDate", result.Parameters[0].Name);
        Assert.Equal("2024-01-01", result.Parameters[0].Value);
    }

    [Fact]
    public void ParameterValueResult_HasValueAndReference()
    {
        // Arrange & Act
        var result = new ParameterValueResult
        {
            Success = true,
            FilePath = "test.xlsx",
            ParameterName = "ReportDate",
            Value = "2024-03-15",
            RefersTo = "Config!B5"
        };

        // Assert
        Assert.Equal("ReportDate", result.ParameterName);
        Assert.Equal("2024-03-15", result.Value);
        Assert.Equal("Config!B5", result.RefersTo);
    }

    [Fact]
    public void WorksheetListResult_WithSheets_HasCorrectStructure()
    {
        // Arrange & Act
        var result = new WorksheetListResult
        {
            Success = true,
            FilePath = "test.xlsx",
            Worksheets = new List<WorksheetInfo>
            {
                new() { Name = "Sheet1", Index = 1, Visible = true },
                new() { Name = "Hidden", Index = 2, Visible = false },
                new() { Name = "Data", Index = 3, Visible = true }
            }
        };

        // Assert
        Assert.Equal(3, result.Worksheets.Count);
        Assert.Equal("Sheet1", result.Worksheets[0].Name);
        Assert.Equal(1, result.Worksheets[0].Index);
        Assert.True(result.Worksheets[0].Visible);
        Assert.False(result.Worksheets[1].Visible);
    }

    [Fact]
    public void WorksheetDataResult_WithData_HasRowsAndColumns()
    {
        // Arrange & Act
        var result = new WorksheetDataResult
        {
            Success = true,
            FilePath = "test.xlsx",
            SheetName = "Data",
            Range = "A1:C3",
            Headers = new List<string> { "Name", "Age", "City" },
            Data = new List<List<object?>>
            {
                new() { "Alice", 30, "NYC" },
                new() { "Bob", 25, "LA" },
                new() { "Charlie", 35, "SF" }
            },
            RowCount = 3,
            ColumnCount = 3
        };

        // Assert
        Assert.Equal(3, result.RowCount);
        Assert.Equal(3, result.ColumnCount);
        Assert.Equal(3, result.Headers.Count);
        Assert.Equal(3, result.Data.Count);
        Assert.Equal("Alice", result.Data[0][0]);
        Assert.Equal(30, result.Data[0][1]);
    }

    [Fact]
    public void ScriptListResult_WithModules_HasCorrectStructure()
    {
        // Arrange & Act
        var result = new ScriptListResult
        {
            Success = true,
            FilePath = "test.xlsm",
            Scripts = new List<ScriptInfo>
            {
                new() 
                { 
                    Name = "Module1", 
                    Type = "Standard",
                    Procedures = new List<string> { "Main", "Helper" },
                    LineCount = 150
                },
                new() 
                { 
                    Name = "Sheet1", 
                    Type = "Worksheet",
                    Procedures = new List<string> { "Worksheet_Change" },
                    LineCount = 45
                }
            }
        };

        // Assert
        Assert.Equal(2, result.Scripts.Count);
        Assert.Equal("Module1", result.Scripts[0].Name);
        Assert.Equal(2, result.Scripts[0].Procedures.Count);
        Assert.Equal(150, result.Scripts[0].LineCount);
    }

    [Fact]
    public void PowerQueryListResult_WithQueries_HasCorrectStructure()
    {
        // Arrange & Act
        var result = new PowerQueryListResult
        {
            Success = true,
            FilePath = "test.xlsx",
            Queries = new List<PowerQueryInfo>
            {
                new() 
                { 
                    Name = "SalesData", 
                    Formula = "let Source = Excel.CurrentWorkbook() in Source",
                    IsConnectionOnly = false
                },
                new() 
                { 
                    Name = "Helper", 
                    Formula = "(x) => x + 1",
                    IsConnectionOnly = true
                }
            }
        };

        // Assert
        Assert.Equal(2, result.Queries.Count);
        Assert.Equal("SalesData", result.Queries[0].Name);
        Assert.False(result.Queries[0].IsConnectionOnly);
        Assert.True(result.Queries[1].IsConnectionOnly);
    }

    [Fact]
    public void PowerQueryViewResult_WithMCode_HasCodeAndMetadata()
    {
        // Arrange & Act
        var result = new PowerQueryViewResult
        {
            Success = true,
            FilePath = "test.xlsx",
            QueryName = "WebData",
            MCode = "let\n    Source = Web.Contents(\"https://api.example.com\")\nin\n    Source",
            CharacterCount = 73,
            IsConnectionOnly = false
        };

        // Assert
        Assert.Equal("WebData", result.QueryName);
        Assert.Contains("Web.Contents", result.MCode);
        Assert.Equal(73, result.CharacterCount);
    }

    [Fact]
    public void VbaTrustResult_Trusted_HasCorrectProperties()
    {
        // Arrange & Act
        var result = new VbaTrustResult
        {
            Success = true,
            IsTrusted = true,
            ComponentCount = 5,
            RegistryPathsSet = new List<string> 
            { 
                @"HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM" 
            },
            ManualInstructions = null
        };

        // Assert
        Assert.True(result.IsTrusted);
        Assert.Equal(5, result.ComponentCount);
        Assert.Single(result.RegistryPathsSet);
        Assert.Null(result.ManualInstructions);
    }

    [Fact]
    public void VbaTrustResult_NotTrusted_HasManualInstructions()
    {
        // Arrange & Act
        var result = new VbaTrustResult
        {
            Success = false,
            IsTrusted = false,
            ComponentCount = 0,
            RegistryPathsSet = new List<string>(),
            ManualInstructions = "Please enable Trust access to VBA project in Excel settings"
        };

        // Assert
        Assert.False(result.IsTrusted);
        Assert.NotNull(result.ManualInstructions);
        Assert.Empty(result.RegistryPathsSet);
    }

    [Fact]
    public void FileValidationResult_ValidFile_HasCorrectProperties()
    {
        // Arrange & Act
        var result = new FileValidationResult
        {
            Success = true,
            FilePath = "test.xlsx",
            Exists = true,
            IsValid = true,
            Extension = ".xlsx",
            Size = 50000
        };

        // Assert
        Assert.True(result.Exists);
        Assert.True(result.IsValid);
        Assert.Equal(".xlsx", result.Extension);
        Assert.Equal(50000, result.Size);
    }

    [Fact]
    public void FileValidationResult_InvalidFile_HasErrorMessage()
    {
        // Arrange & Act
        var result = new FileValidationResult
        {
            Success = false,
            FilePath = "test.txt",
            Exists = true,
            IsValid = false,
            Extension = ".txt",
            Size = 100,
            ErrorMessage = "Not a valid Excel file extension"
        };

        // Assert
        Assert.False(result.IsValid);
        Assert.Equal(".txt", result.Extension);
        Assert.NotNull(result.ErrorMessage);
    }
}
