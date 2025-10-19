using ExcelMcp.Core.Commands;
using ModelContextProtocol.Server;
using System.ComponentModel;
using System.Text.Json;

namespace ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel automation tools for Model Context Protocol (MCP) server.
/// Provides 6 resource-based tools for comprehensive Excel operations.
/// </summary>
[McpServerToolType]
public static class ExcelTools
{
    #region File Operations

    /// <summary>
    /// Manage Excel files - create, validate, and check file operations
    /// </summary>
    [McpServerTool(Name = "excel_file")]
    [Description("Create, validate, and manage Excel files (.xlsx, .xlsm). Supports actions: create-empty, validate, check-exists.")]
    public static string ExcelFile(
        [Description("Action to perform: create-empty, validate, check-exists")] string action,
        [Description("Excel file path (.xlsx or .xlsm extension)")] string filePath,
        [Description("Optional: macro-enabled flag for create-empty (default: false)")] bool macroEnabled = false)
    {
        try
        {
            var fileCommands = new FileCommands();

            return action.ToLowerInvariant() switch
            {
                "create-empty" => CreateEmptyFile(fileCommands, filePath, macroEnabled),
                "validate" => ValidateFile(filePath),
                "check-exists" => CheckFileExists(filePath),
                _ => JsonSerializer.Serialize(new { error = $"Unknown action '{action}'. Supported: create-empty, validate, check-exists" })
            };
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                error = ex.Message,
                action,
                filePath
            });
        }
    }

    private static string CreateEmptyFile(FileCommands fileCommands, string filePath, bool macroEnabled)
    {
        var extension = macroEnabled ? ".xlsm" : ".xlsx";
        if (!filePath.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
        {
            filePath = Path.ChangeExtension(filePath, extension);
        }

        var result = fileCommands.CreateEmpty(new[] { "create-empty", filePath });
        if (result == 0)
        {
            return JsonSerializer.Serialize(new
            {
                success = true,
                filePath,
                macroEnabled,
                message = "Excel file created successfully"
            });
        }
        else
        {
            return JsonSerializer.Serialize(new
            {
                error = "Failed to create Excel file",
                filePath
            });
        }
    }

    private static string ValidateFile(string filePath)
    {
        if (!File.Exists(filePath))
        {
            return JsonSerializer.Serialize(new
            {
                valid = false,
                error = "File does not exist",
                filePath
            });
        }

        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        if (extension != ".xlsx" && extension != ".xlsm")
        {
            return JsonSerializer.Serialize(new
            {
                valid = false,
                error = "Invalid file extension. Expected .xlsx or .xlsm",
                filePath
            });
        }

        return JsonSerializer.Serialize(new
        {
            valid = true,
            filePath,
            extension
        });
    }

    private static string CheckFileExists(string filePath)
    {
        var exists = File.Exists(filePath);
        var size = exists ? new FileInfo(filePath).Length : 0;
        return JsonSerializer.Serialize(new
        {
            exists,
            filePath,
            size
        });
    }

    #endregion

    #region Power Query Operations

    /// <summary>
    /// Manage Power Query M code and data connections
    /// </summary>
    [McpServerTool(Name = "excel_powerquery")]
    [Description("Manage Power Query M code, connections, and data transformations. Actions: list, view, import, export, update, refresh, loadto, delete.")]
    public static string ExcelPowerQuery(
        [Description("Action to perform: list, view, import, export, update, refresh, loadto, delete")] string action,
        [Description("Excel file path")] string filePath,
        [Description("Power Query name (required for: view, import, export, update, refresh, loadto, delete)")] string? queryName = null,
        [Description("Source file path for import/update operations or target file for export")] string? sourceOrTargetPath = null,
        [Description("Target worksheet name for loadto action")] string? targetSheet = null,
        [Description("M code content for update operations")] string? mCode = null)
    {
        try
        {
            var powerQueryCommands = new PowerQueryCommands();

            return action.ToLowerInvariant() switch
            {
                "list" => ExecutePowerQueryCommand(powerQueryCommands, "List", filePath),
                "view" => ExecutePowerQueryCommand(powerQueryCommands, "View", filePath, queryName),
                "import" => ExecutePowerQueryCommand(powerQueryCommands, "Import", filePath, queryName, sourceOrTargetPath),
                "export" => ExecutePowerQueryCommand(powerQueryCommands, "Export", filePath, queryName, sourceOrTargetPath),
                "update" => ExecutePowerQueryCommand(powerQueryCommands, "Update", filePath, queryName, sourceOrTargetPath),
                "refresh" => ExecutePowerQueryCommand(powerQueryCommands, "Refresh", filePath, queryName),
                "loadto" => ExecutePowerQueryCommand(powerQueryCommands, "LoadTo", filePath, queryName, targetSheet),
                "delete" => ExecutePowerQueryCommand(powerQueryCommands, "Delete", filePath, queryName),
                _ => JsonSerializer.Serialize(new { error = $"Unknown action '{action}'. Supported: list, view, import, export, update, refresh, loadto, delete" })
            };
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                error = ex.Message,
                action,
                filePath,
                queryName
            });
        }
    }

    private static string ExecutePowerQueryCommand(PowerQueryCommands commands, string method, string filePath, string? arg1 = null, string? arg2 = null)
    {
        var args = new List<string> { $"pq-{method.ToLowerInvariant()}", filePath };
        if (!string.IsNullOrEmpty(arg1)) args.Add(arg1);
        if (!string.IsNullOrEmpty(arg2)) args.Add(arg2);

        var methodInfo = typeof(PowerQueryCommands).GetMethod(method);
        if (methodInfo == null)
        {
            return JsonSerializer.Serialize(new { error = $"Method {method} not found" });
        }

        var result = (int)methodInfo.Invoke(commands, new object[] { args.ToArray() })!;
        if (result == 0)
        {
            return JsonSerializer.Serialize(new
            {
                success = true,
                action = method.ToLowerInvariant(),
                filePath
            });
        }
        else
        {
            return JsonSerializer.Serialize(new
            {
                error = "Operation failed",
                action = method.ToLowerInvariant(),
                filePath
            });
        }
    }

    #endregion

    #region Worksheet Operations

    /// <summary>
    /// CRUD operations on worksheets and cell ranges
    /// </summary>
    [McpServerTool(Name = "excel_worksheet")]
    [Description("Manage worksheets and data ranges. Actions: list, read, write, create, rename, copy, delete, clear, append.")]
    public static string ExcelWorksheet(
        [Description("Action to perform: list, read, write, create, rename, copy, delete, clear, append")] string action,
        [Description("Excel file path")] string filePath,
        [Description("Worksheet name (required for most actions)")] string? sheetName = null,
        [Description("Cell range (e.g., 'A1:D10') or CSV file path for data operations")] string? rangeOrDataPath = null,
        [Description("Target name for rename/copy operations")] string? targetName = null)
    {
        try
        {
            var sheetCommands = new SheetCommands();

            return action.ToLowerInvariant() switch
            {
                "list" => ExecuteSheetCommand(sheetCommands, "List", filePath),
                "read" => ExecuteSheetCommand(sheetCommands, "Read", filePath, sheetName, rangeOrDataPath),
                "write" => ExecuteSheetCommand(sheetCommands, "Write", filePath, sheetName, rangeOrDataPath),
                "create" => ExecuteSheetCommand(sheetCommands, "Create", filePath, sheetName),
                "rename" => ExecuteSheetCommand(sheetCommands, "Rename", filePath, sheetName, targetName),
                "copy" => ExecuteSheetCommand(sheetCommands, "Copy", filePath, sheetName, targetName),
                "delete" => ExecuteSheetCommand(sheetCommands, "Delete", filePath, sheetName),
                "clear" => ExecuteSheetCommand(sheetCommands, "Clear", filePath, sheetName, rangeOrDataPath),
                "append" => ExecuteSheetCommand(sheetCommands, "Append", filePath, sheetName, rangeOrDataPath),
                _ => JsonSerializer.Serialize(new { error = $"Unknown action '{action}'. Supported: list, read, write, create, rename, copy, delete, clear, append" })
            };
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                error = ex.Message,
                action,
                filePath,
                sheetName
            });
        }
    }

    private static string ExecuteSheetCommand(SheetCommands commands, string method, string filePath, string? arg1 = null, string? arg2 = null)
    {
        var args = new List<string> { $"sheet-{method.ToLowerInvariant()}", filePath };
        if (!string.IsNullOrEmpty(arg1)) args.Add(arg1);
        if (!string.IsNullOrEmpty(arg2)) args.Add(arg2);

        var methodInfo = typeof(SheetCommands).GetMethod(method);
        if (methodInfo == null)
        {
            return JsonSerializer.Serialize(new { error = $"Method {method} not found" });
        }

        var result = (int)methodInfo.Invoke(commands, new object[] { args.ToArray() })!;
        if (result == 0)
        {
            return JsonSerializer.Serialize(new
            {
                success = true,
                action = method.ToLowerInvariant(),
                filePath
            });
        }
        else
        {
            return JsonSerializer.Serialize(new
            {
                error = "Operation failed",
                action = method.ToLowerInvariant(),
                filePath
            });
        }
    }

    #endregion

    #region Parameter Operations

    /// <summary>
    /// Manage Excel named ranges as parameters
    /// </summary>
    [McpServerTool(Name = "excel_parameter")]
    [Description("Manage named ranges as parameters for configuration. Actions: list, get, set, create, delete.")]
    public static string ExcelParameter(
        [Description("Action to perform: list, get, set, create, delete")] string action,
        [Description("Excel file path")] string filePath,
        [Description("Parameter/named range name (required for: get, set, create, delete)")] string? paramName = null,
        [Description("Parameter value for set operations or cell reference for create (e.g., 'Sheet1!A1')")] string? valueOrReference = null)
    {
        try
        {
            var paramCommands = new ParameterCommands();

            return action.ToLowerInvariant() switch
            {
                "list" => ExecuteParameterCommand(paramCommands, "List", filePath),
                "get" => ExecuteParameterCommand(paramCommands, "Get", filePath, paramName),
                "set" => ExecuteParameterCommand(paramCommands, "Set", filePath, paramName, valueOrReference),
                "create" => ExecuteParameterCommand(paramCommands, "Create", filePath, paramName, valueOrReference),
                "delete" => ExecuteParameterCommand(paramCommands, "Delete", filePath, paramName),
                _ => JsonSerializer.Serialize(new { error = $"Unknown action '{action}'. Supported: list, get, set, create, delete" })
            };
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                error = ex.Message,
                action,
                filePath,
                paramName
            });
        }
    }

    private static string ExecuteParameterCommand(ParameterCommands commands, string method, string filePath, string? arg1 = null, string? arg2 = null)
    {
        var args = new List<string> { $"param-{method.ToLowerInvariant()}", filePath };
        if (!string.IsNullOrEmpty(arg1)) args.Add(arg1);
        if (!string.IsNullOrEmpty(arg2)) args.Add(arg2);

        var methodInfo = typeof(ParameterCommands).GetMethod(method);
        if (methodInfo == null)
        {
            return JsonSerializer.Serialize(new { error = $"Method {method} not found" });
        }

        var result = (int)methodInfo.Invoke(commands, new object[] { args.ToArray() })!;
        if (result == 0)
        {
            return JsonSerializer.Serialize(new
            {
                success = true,
                action = method.ToLowerInvariant(),
                filePath
            });
        }
        else
        {
            return JsonSerializer.Serialize(new
            {
                error = "Operation failed",
                action = method.ToLowerInvariant(),
                filePath
            });
        }
    }

    #endregion

    #region Cell Operations

    /// <summary>
    /// Individual cell operations for values and formulas
    /// </summary>
    [McpServerTool(Name = "excel_cell")]
    [Description("Get/set individual cell values and formulas. Actions: get-value, set-value, get-formula, set-formula.")]
    public static string ExcelCell(
        [Description("Action to perform: get-value, set-value, get-formula, set-formula")] string action,
        [Description("Excel file path")] string filePath,
        [Description("Worksheet name")] string sheetName,
        [Description("Cell address (e.g., 'A1', 'B5')")] string cellAddress,
        [Description("Value or formula to set (required for set operations)")] string? valueOrFormula = null)
    {
        try
        {
            var cellCommands = new CellCommands();

            return action.ToLowerInvariant() switch
            {
                "get-value" => ExecuteCellCommand(cellCommands, "GetValue", filePath, sheetName, cellAddress),
                "set-value" => ExecuteCellCommand(cellCommands, "SetValue", filePath, sheetName, cellAddress, valueOrFormula),
                "get-formula" => ExecuteCellCommand(cellCommands, "GetFormula", filePath, sheetName, cellAddress),
                "set-formula" => ExecuteCellCommand(cellCommands, "SetFormula", filePath, sheetName, cellAddress, valueOrFormula),
                _ => JsonSerializer.Serialize(new { error = $"Unknown action '{action}'. Supported: get-value, set-value, get-formula, set-formula" })
            };
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                error = ex.Message,
                action,
                filePath,
                sheetName,
                cellAddress
            });
        }
    }

    private static string ExecuteCellCommand(CellCommands commands, string method, string filePath, string sheetName, string cellAddress, string? valueOrFormula = null)
    {
        var args = new List<string> { $"cell-{method.ToKebabCase()}", filePath, sheetName, cellAddress };
        if (!string.IsNullOrEmpty(valueOrFormula)) args.Add(valueOrFormula);

        var methodInfo = typeof(CellCommands).GetMethod(method);
        if (methodInfo == null)
        {
            return JsonSerializer.Serialize(new { error = $"Method {method} not found" });
        }

        var result = (int)methodInfo.Invoke(commands, new object[] { args.ToArray() })!;
        if (result == 0)
        {
            return JsonSerializer.Serialize(new
            {
                success = true,
                action = method.ToKebabCase(),
                filePath,
                sheetName,
                cellAddress
            });
        }
        else
        {
            return JsonSerializer.Serialize(new
            {
                error = "Operation failed",
                action = method.ToKebabCase(),
                filePath
            });
        }
    }

    #endregion

    #region VBA Script Operations

    /// <summary>
    /// VBA script management and execution (requires .xlsm files)
    /// </summary>
    [McpServerTool(Name = "excel_vba")]
    [Description("Manage and execute VBA scripts (.xlsm files only). Actions: list, export, import, update, run, delete, setup-trust, check-trust.")]
    public static string ExcelVba(
        [Description("Action to perform: list, export, import, update, run, delete, setup-trust, check-trust")] string action,
        [Description("Excel file path (.xlsm required for most operations)")] string? filePath = null,
        [Description("VBA module name (required for: export, import, update, delete)")] string? moduleName = null,
        [Description("VBA file path for import/export or procedure name for run")] string? vbaFileOrProcedure = null,
        [Description("Parameters for VBA procedure execution (space-separated)")] string? parameters = null)
    {
        try
        {
            var scriptCommands = new ScriptCommands();
            var setupCommands = new SetupCommands();

            return action.ToLowerInvariant() switch
            {
                "setup-trust" => ExecuteSetupCommand(setupCommands, "SetupVbaTrust"),
                "check-trust" => ExecuteSetupCommand(setupCommands, "CheckVbaTrust"),
                "list" => ExecuteScriptCommand(scriptCommands, "List", filePath!),
                "export" => ExecuteScriptCommand(scriptCommands, "Export", filePath!, moduleName, vbaFileOrProcedure),
                "import" => ExecuteScriptCommand(scriptCommands, "Import", filePath!, moduleName, vbaFileOrProcedure),
                "update" => ExecuteScriptCommand(scriptCommands, "Update", filePath!, moduleName, vbaFileOrProcedure),
                "run" => ExecuteScriptRunCommand(scriptCommands, filePath!, vbaFileOrProcedure, parameters),
                "delete" => ExecuteScriptCommand(scriptCommands, "Delete", filePath!, moduleName),
                _ => JsonSerializer.Serialize(new { error = $"Unknown action '{action}'. Supported: list, export, import, update, run, delete, setup-trust, check-trust" })
            };
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                error = ex.Message,
                action,
                filePath,
                moduleName
            });
        }
    }

    private static string ExecuteSetupCommand(SetupCommands commands, string method)
    {
        var args = new[] { method.ToKebabCase() };
        var methodInfo = typeof(SetupCommands).GetMethod(method);
        if (methodInfo == null)
        {
            return JsonSerializer.Serialize(new { error = $"Method {method} not found" });
        }

        var result = (int)methodInfo.Invoke(commands, new object[] { args })!;
        if (result == 0)
        {
            return JsonSerializer.Serialize(new
            {
                success = true,
                action = method.ToKebabCase()
            });
        }
        else
        {
            return JsonSerializer.Serialize(new
            {
                error = "Operation failed",
                action = method.ToKebabCase()
            });
        }
    }

    private static string ExecuteScriptCommand(ScriptCommands commands, string method, string filePath, string? arg1 = null, string? arg2 = null)
    {
        var args = new List<string> { $"script-{method.ToLowerInvariant()}", filePath };
        if (!string.IsNullOrEmpty(arg1)) args.Add(arg1);
        if (!string.IsNullOrEmpty(arg2)) args.Add(arg2);

        var methodInfo = typeof(ScriptCommands).GetMethod(method);
        if (methodInfo == null)
        {
            return JsonSerializer.Serialize(new { error = $"Method {method} not found" });
        }

        var result = (int)methodInfo.Invoke(commands, new object[] { args.ToArray() })!;
        if (result == 0)
        {
            return JsonSerializer.Serialize(new
            {
                success = true,
                action = method.ToLowerInvariant(),
                filePath
            });
        }
        else
        {
            return JsonSerializer.Serialize(new
            {
                error = "Operation failed",
                action = method.ToLowerInvariant(),
                filePath
            });
        }
    }

    private static string ExecuteScriptRunCommand(ScriptCommands commands, string filePath, string? procedureName, string? parameters)
    {
        var args = new List<string> { "script-run", filePath };
        if (!string.IsNullOrEmpty(procedureName)) args.Add(procedureName);
        if (!string.IsNullOrEmpty(parameters))
        {
            // Split parameters by space and add each as separate argument
            args.AddRange(parameters.Split(' ', StringSplitOptions.RemoveEmptyEntries));
        }

        var result = commands.Run(args.ToArray());
        if (result == 0)
        {
            return JsonSerializer.Serialize(new
            {
                success = true,
                action = "run",
                filePath,
                procedure = procedureName
            });
        }
        else
        {
            return JsonSerializer.Serialize(new
            {
                error = "Operation failed",
                action = "run",
                filePath
            });
        }
    }

    #endregion
}

/// <summary>
/// Extension methods for string formatting
/// </summary>
public static class StringExtensions
{
    public static string ToKebabCase(this string text)
    {
        if (string.IsNullOrEmpty(text)) return text;

        var result = new System.Text.StringBuilder();
        for (int i = 0; i < text.Length; i++)
        {
            if (i > 0 && char.IsUpper(text[i]))
            {
                result.Append('-');
            }
            result.Append(char.ToLowerInvariant(text[i]));
        }
        return result.ToString();
    }
}