using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit.Serialization;

/// <summary>
/// MCP Server-specific tests for JSON serialization of Result objects - Unit tests, no Excel required
/// 
/// LAYER RESPONSIBILITY:
/// - ✅ Test JSON serialization of all Result types
/// - ✅ Test property naming (camelCase for MCP protocol)
/// - ✅ Test null value handling in JSON
/// - ✅ Test deserialization roundtrip
/// - ❌ DO NOT test Excel operations (that's Core's responsibility)
/// - ❌ DO NOT test CLI output formatting (that's CLI's responsibility)
/// 
/// These tests verify that MCP Server correctly serializes Core Result objects to JSON for MCP protocol responses.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
public class ResultSerializationTests
{
    private readonly JsonSerializerOptions _options = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false
    };

    [Fact]
    public void OperationResult_Success_SerializesToJson()
    {
        // Arrange
        var result = new OperationResult
        {
            Success = true,
            FilePath = "test.xlsx",
            Action = "create",
            ErrorMessage = null
        };

        // Act
        var json = JsonSerializer.Serialize(result, _options);
        var deserialized = JsonSerializer.Deserialize<OperationResult>(json, _options);

        // Assert
        Assert.NotNull(json);
        Assert.Contains("\"success\":true", json);
        Assert.Contains("\"action\":\"create\"", json);
        Assert.NotNull(deserialized);
        Assert.True(deserialized.Success);
        Assert.Equal("create", deserialized.Action);
    }

    [Fact]
    public void OperationResult_Failure_SerializesErrorMessage()
    {
        // Arrange
        var result = new OperationResult
        {
            Success = false,
            FilePath = "test.xlsx",
            Action = "delete",
            ErrorMessage = "File not found"
        };

        // Act
        var json = JsonSerializer.Serialize(result, _options);
        var deserialized = JsonSerializer.Deserialize<OperationResult>(json, _options);

        // Assert
        Assert.Contains("\"success\":false", json);
        Assert.Contains("\"errorMessage\":\"File not found\"", json);
        Assert.NotNull(deserialized);
        Assert.False(deserialized.Success);
        Assert.Equal("File not found", deserialized.ErrorMessage);
    }

    [Fact]
    public void CellValueResult_WithData_SerializesToJson()
    {
        // Arrange
        var result = new CellValueResult
        {
            Success = true,
            FilePath = "test.xlsx",
            CellAddress = "A1",
            Value = "Hello World",
            ValueType = "String",
            Formula = null
        };

        // Act
        var json = JsonSerializer.Serialize(result, _options);
        var deserialized = JsonSerializer.Deserialize<CellValueResult>(json, _options);

        // Assert
        Assert.Contains("\"cellAddress\":\"A1\"", json);
        Assert.Contains("\"value\":\"Hello World\"", json);
        Assert.NotNull(deserialized);
        Assert.Equal("A1", deserialized.CellAddress);
        Assert.Equal("Hello World", deserialized.Value?.ToString());
    }

    [Fact]
    public void WorksheetListResult_WithSheets_SerializesToJson()
    {
        // Arrange
        var result = new WorksheetListResult
        {
            Success = true,
            FilePath = "test.xlsx",
            Worksheets = new List<WorksheetInfo>
            {
                new() { Name = "Sheet1", Index = 1, Visible = true },
                new() { Name = "Sheet2", Index = 2, Visible = false }
            }
        };

        // Act
        var json = JsonSerializer.Serialize(result, _options);
        var deserialized = JsonSerializer.Deserialize<WorksheetListResult>(json, _options);

        // Assert
        Assert.Contains("\"worksheets\":", json);
        Assert.Contains("\"Sheet1\"", json);
        Assert.Contains("\"Sheet2\"", json);
        Assert.NotNull(deserialized);
        Assert.Equal(2, deserialized.Worksheets.Count);
        Assert.Equal("Sheet1", deserialized.Worksheets[0].Name);
    }

    [Fact]
    public void WorksheetDataResult_WithData_SerializesToJson()
    {
        // Arrange
        var result = new WorksheetDataResult
        {
            Success = true,
            FilePath = "test.xlsx",
            SheetName = "Data",
            Range = "A1:B2",
            Headers = new List<string> { "Name", "Age" },
            Data = new List<List<object?>>
            {
                new() { "Alice", 30 },
                new() { "Bob", 25 }
            },
            RowCount = 2,
            ColumnCount = 2
        };

        // Act
        var json = JsonSerializer.Serialize(result, _options);
        var deserialized = JsonSerializer.Deserialize<WorksheetDataResult>(json, _options);

        // Assert
        Assert.Contains("\"headers\":", json);
        Assert.Contains("\"data\":", json);
        Assert.Contains("\"Alice\"", json);
        Assert.NotNull(deserialized);
        Assert.Equal(2, deserialized.Headers.Count);
        Assert.Equal(2, deserialized.Data.Count);
    }

    [Fact]
    public void ParameterListResult_WithParameters_SerializesToJson()
    {
        // Arrange
        var result = new ParameterListResult
        {
            Success = true,
            FilePath = "test.xlsx",
            Parameters = new List<ParameterInfo>
            {
                new() { Name = "StartDate", Value = "2024-01-01", RefersTo = "Config!A1" },
                new() { Name = "EndDate", Value = "2024-12-31", RefersTo = "Config!A2" }
            }
        };

        // Act
        var json = JsonSerializer.Serialize(result, _options);
        var deserialized = JsonSerializer.Deserialize<ParameterListResult>(json, _options);

        // Assert
        Assert.Contains("\"parameters\":", json);
        Assert.Contains("\"StartDate\"", json);
        Assert.Contains("\"EndDate\"", json);
        Assert.NotNull(deserialized);
        Assert.Equal(2, deserialized.Parameters.Count);
    }

    [Fact]
    public void ScriptListResult_WithModules_SerializesToJson()
    {
        // Arrange
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
                    LineCount = 150,
                    Procedures = new List<string> { "Main", "Helper" }
                }
            }
        };

        // Act
        var json = JsonSerializer.Serialize(result, _options);
        var deserialized = JsonSerializer.Deserialize<ScriptListResult>(json, _options);

        // Assert
        Assert.Contains("\"scripts\":", json);
        Assert.Contains("\"Module1\"", json);
        Assert.Contains("\"procedures\":", json);
        Assert.NotNull(deserialized);
        Assert.Single(deserialized.Scripts);
        Assert.Equal(150, deserialized.Scripts[0].LineCount);
    }

    [Fact]
    public void PowerQueryListResult_WithQueries_SerializesToJson()
    {
        // Arrange
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
                }
            }
        };

        // Act
        var json = JsonSerializer.Serialize(result, _options);
        var deserialized = JsonSerializer.Deserialize<PowerQueryListResult>(json, _options);

        // Assert
        Assert.Contains("\"queries\":", json);
        Assert.Contains("\"SalesData\"", json);
        Assert.Contains("\"isConnectionOnly\"", json);
        Assert.NotNull(deserialized);
        Assert.Single(deserialized.Queries);
    }

    [Fact]
    public void PowerQueryViewResult_WithMCode_SerializesToJson()
    {
        // Arrange
        var result = new PowerQueryViewResult
        {
            Success = true,
            FilePath = "test.xlsx",
            QueryName = "WebData",
            MCode = "let\n    Source = Web.Contents(\"https://api.example.com\")\nin\n    Source",
            CharacterCount = 73,
            IsConnectionOnly = false
        };

        // Act
        var json = JsonSerializer.Serialize(result, _options);
        var deserialized = JsonSerializer.Deserialize<PowerQueryViewResult>(json, _options);

        // Assert
        Assert.Contains("\"queryName\":\"WebData\"", json);
        Assert.Contains("\"mCode\":", json);
        Assert.Contains("Web.Contents", json);
        Assert.NotNull(deserialized);
        Assert.Equal("WebData", deserialized.QueryName);
        Assert.Equal(73, deserialized.CharacterCount);
    }

    [Fact]
    public void VbaTrustResult_SerializesToJson()
    {
        // Arrange
        var result = new VbaTrustResult
        {
            Success = true,
            IsTrusted = true,
            ComponentCount = 5,
            RegistryPathsSet = new List<string> { @"HKCU\Software\Microsoft\Office\16.0" },
            ManualInstructions = null
        };

        // Act
        var json = JsonSerializer.Serialize(result, _options);
        var deserialized = JsonSerializer.Deserialize<VbaTrustResult>(json, _options);

        // Assert
        Assert.Contains("\"isTrusted\":true", json);
        Assert.Contains("\"componentCount\":5", json);
        Assert.NotNull(deserialized);
        Assert.True(deserialized.IsTrusted);
        Assert.Equal(5, deserialized.ComponentCount);
    }

    [Fact]
    public void FileValidationResult_SerializesToJson()
    {
        // Arrange
        var result = new FileValidationResult
        {
            Success = true,
            FilePath = "test.xlsx",
            Exists = true,
            IsValid = true,
            Extension = ".xlsx",
            Size = 50000
        };

        // Act
        var json = JsonSerializer.Serialize(result, _options);
        var deserialized = JsonSerializer.Deserialize<FileValidationResult>(json, _options);

        // Assert
        Assert.Contains("\"exists\":true", json);
        Assert.Contains("\"isValid\":true", json);
        Assert.Contains("\"extension\":\".xlsx\"", json);
        Assert.NotNull(deserialized);
        Assert.True(deserialized.Exists);
        Assert.Equal(".xlsx", deserialized.Extension);
    }

    [Fact]
    public void NullValues_SerializeCorrectly()
    {
        // Arrange
        var result = new OperationResult
        {
            Success = true,
            FilePath = "test.xlsx",
            Action = "create",
            ErrorMessage = null
        };

        // Act
        var json = JsonSerializer.Serialize(result, _options);

        // Assert
        // Null values should be included in JSON (MCP Server needs complete responses)
        Assert.Contains("\"errorMessage\":null", json);
    }

    [Fact]
    public void EmptyCollections_SerializeAsEmptyArrays()
    {
        // Arrange
        var result = new WorksheetListResult
        {
            Success = true,
            FilePath = "test.xlsx",
            Worksheets = new List<WorksheetInfo>()
        };

        // Act
        var json = JsonSerializer.Serialize(result, _options);
        var deserialized = JsonSerializer.Deserialize<WorksheetListResult>(json, _options);

        // Assert
        Assert.Contains("\"worksheets\":[]", json);
        Assert.NotNull(deserialized);
        Assert.Empty(deserialized.Worksheets);
    }

    [Fact]
    public void ComplexNestedData_SerializesCorrectly()
    {
        // Arrange
        var result = new WorksheetDataResult
        {
            Success = true,
            FilePath = "test.xlsx",
            SheetName = "Complex",
            Range = "A1:C2",
            Headers = new List<string> { "String", "Number", "Boolean" },
            Data = new List<List<object?>>
            {
                new() { "text", 42, true },
                new() { null, 3.14, false }
            },
            RowCount = 2,
            ColumnCount = 3
        };

        // Act
        var json = JsonSerializer.Serialize(result, _options);
        var deserialized = JsonSerializer.Deserialize<WorksheetDataResult>(json, _options);

        // Assert
        Assert.Contains("\"String\"", json);
        Assert.Contains("\"Number\"", json);
        Assert.Contains("\"Boolean\"", json);
        Assert.Contains("42", json);
        Assert.Contains("3.14", json);
        Assert.Contains("true", json);
        Assert.Contains("false", json);
        Assert.NotNull(deserialized);
        Assert.Equal(2, deserialized.Data.Count);
        Assert.Null(deserialized.Data[1][0]); // Null value in data
    }
}
