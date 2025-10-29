using System.Text.RegularExpressions;

#pragma warning disable CS1591 // Missing XML comment - Description property documents each action

namespace Sbroenne.ExcelMcp.Core.Models.Validation;

/// <summary>
/// Central registry of all action definitions
/// Domain-focused definitions without client-specific concerns (CLI/MCP)
/// </summary>
public static class ActionDefinitions
{
    /// <summary>
    /// Shared parameter definitions used across multiple domains
    /// </summary>
    private static class SharedParams
    {
        public static readonly ParameterDefinition ExcelPath = new()
        {
            Name = "excelPath",
            Required = true,
            FileExtensions = new[] { "xlsx", "xlsm" },
            Description = "Excel file path (.xlsx or .xlsm)"
        };
    }

    /// <summary>
    /// Power Query action definitions
    /// </summary>
    public static class PowerQuery
    {
        private const string Domain = "PowerQuery";

        /// <summary>
        /// Common parameters used across multiple actions
        /// </summary>
        private static class CommonParams
        {
            public static readonly ParameterDefinition ExcelPath = SharedParams.ExcelPath;

            public static readonly ParameterDefinition QueryName = new()
            {
                Name = "queryName",
                Required = true,
                MaxLength = 255,
                MinLength = 1,
                Description = "Power Query name"
            };

            public static readonly ParameterDefinition SourcePath = new()
            {
                Name = "sourcePath",
                Required = true,
                FileExtensions = new[] { "pq", "txt", "m" },
                Description = "Source .pq file path"
            };

            public static readonly ParameterDefinition TargetPath = new()
            {
                Name = "targetPath",
                Required = true,
                FileExtensions = new[] { "pq", "txt", "m" },
                Description = "Target file path for export"
            };

            public static readonly ParameterDefinition TargetSheet = new()
            {
                Name = "targetSheet",
                Required = false,
                MaxLength = 31,
                MinLength = 1,
                Pattern = @"^[^[\]/*?\\:]+$",
                Description = "Target worksheet name"
            };

            public static readonly ParameterDefinition PrivacyLevel = new()
            {
                Name = "privacyLevel",
                Required = false,
                AllowedValues = new[] { "None", "Private", "Organizational", "Public" },
                Description = "Privacy level for data combining"
            };
        }

        public static readonly ActionDefinition List = new()
        {
            Domain = Domain,
            Name = "list",
            Parameters = new[] { CommonParams.ExcelPath },
            Description = "List all Power Queries in workbook"
        };

        public static readonly ActionDefinition View = new()
        {
            Domain = Domain,
            Name = "view",
            Parameters = new[] { CommonParams.ExcelPath, CommonParams.QueryName },
            Description = "View Power Query M code"
        };

        public static readonly ActionDefinition Import = new()
        {
            Domain = Domain,
            Name = "import",
            Parameters = new[] 
            { 
                CommonParams.ExcelPath, 
                CommonParams.QueryName, 
                CommonParams.SourcePath,
                CommonParams.PrivacyLevel 
            },
            Description = "Import Power Query from .pq file"
        };

        public static readonly ActionDefinition Export = new()
        {
            Domain = Domain,
            Name = "export",
            Parameters = new[]
            {
                CommonParams.ExcelPath,
                CommonParams.QueryName,
                CommonParams.TargetPath
            },
            Description = "Export Power Query M code to file"
        };

        public static readonly ActionDefinition Update = new()
        {
            Domain = Domain,
            Name = "update",
            Parameters = new[]
            {
                CommonParams.ExcelPath,
                CommonParams.QueryName,
                CommonParams.SourcePath,
                CommonParams.PrivacyLevel
            },
            Description = "Update Power Query M code from file"
        };

        public static readonly ActionDefinition Delete = new()
        {
            Domain = Domain,
            Name = "delete",
            Parameters = new[]
            {
                CommonParams.ExcelPath,
                CommonParams.QueryName
            },
            Description = "Delete Power Query"
        };

        public static readonly ActionDefinition Refresh = new()
        {
            Domain = Domain,
            Name = "refresh",
            Parameters = new[]
            {
                CommonParams.ExcelPath,
                CommonParams.QueryName
            },
            Description = "Refresh Power Query data from source"
        };

        public static readonly ActionDefinition SetLoadToTable = new()
        {
            Domain = Domain,
            Name = "set-load-to-table",
            Parameters = new[]
            {
                CommonParams.ExcelPath,
                CommonParams.QueryName,
                CommonParams.TargetSheet,
                CommonParams.PrivacyLevel
            },
            Description = "Configure query to load data to worksheet table"
        };

        public static readonly ActionDefinition SetLoadToDataModel = new()
        {
            Domain = Domain,
            Name = "set-load-to-data-model",
            Parameters = new[]
            {
                CommonParams.ExcelPath,
                CommonParams.QueryName,
                CommonParams.PrivacyLevel
            },
            Description = "Configure query to load data to data model"
        };

        public static readonly ActionDefinition SetLoadToBoth = new()
        {
            Domain = Domain,
            Name = "set-load-to-both",
            Parameters = new[]
            {
                CommonParams.ExcelPath,
                CommonParams.QueryName,
                CommonParams.TargetSheet,
                CommonParams.PrivacyLevel
            },
            Description = "Configure query to load to both table and data model"
        };

        public static readonly ActionDefinition SetConnectionOnly = new()
        {
            Domain = Domain,
            Name = "set-connection-only",
            Parameters = new[]
            {
                CommonParams.ExcelPath,
                CommonParams.QueryName
            },
            Description = "Configure query as connection-only (no data loading)"
        };

        public static readonly ActionDefinition GetLoadConfig = new()
        {
            Domain = Domain,
            Name = "get-load-config",
            Parameters = new[]
            {
                CommonParams.ExcelPath,
                CommonParams.QueryName
            },
            Description = "Get current load configuration for query"
        };

        /// <summary>
        /// Gets all PowerQuery action definitions
        /// </summary>
        public static IEnumerable<ActionDefinition> GetAll()
        {
            yield return List;
            yield return View;
            yield return Import;
            yield return Export;
            yield return Update;
            yield return Delete;
            yield return Refresh;
            yield return SetLoadToTable;
            yield return SetLoadToDataModel;
            yield return SetLoadToBoth;
            yield return SetConnectionOnly;
            yield return GetLoadConfig;
        }

        /// <summary>
        /// Gets action definition by name
        /// </summary>
        public static ActionDefinition? GetByName(string actionName)
        {
            return GetAll().FirstOrDefault(a => 
                string.Equals(a.Name, actionName, StringComparison.OrdinalIgnoreCase));
        }
    }

    /// <summary>
    /// Parameter (Named Range) action definitions
    /// </summary>
    public static class Parameter
    {
        private const string Domain = "Parameter";

        private static class CommonParams
        {
            public static readonly ParameterDefinition ExcelPath = SharedParams.ExcelPath;

            public static readonly ParameterDefinition ParameterName = new()
            {
                Name = "parameterName",
                Required = true,
                MaxLength = 255,
                MinLength = 1,
                Description = "Named range parameter name"
            };

            public static readonly ParameterDefinition Value = new()
            {
                Name = "value",
                Required = true,
                Description = "Parameter value"
            };

            public static readonly ParameterDefinition Reference = new()
            {
                Name = "reference",
                Required = true,
                Pattern = @"^=?[A-Za-z]+[0-9]+$|^=?[A-Za-z]+[0-9]+:[A-Za-z]+[0-9]+$",
                Description = "Cell reference (e.g., A1 or A1:B2)"
            };
        }

        public static readonly ActionDefinition List = new()
        {
            Domain = Domain,
            Name = "list",
            Parameters = new[] { CommonParams.ExcelPath },
            Description = "List all named range parameters"
        };

        public static readonly ActionDefinition Get = new()
        {
            Domain = Domain,
            Name = "get",
            Parameters = new[] { CommonParams.ExcelPath, CommonParams.ParameterName },
            Description = "Get parameter value"
        };

        public static readonly ActionDefinition Set = new()
        {
            Domain = Domain,
            Name = "set",
            Parameters = new[] { CommonParams.ExcelPath, CommonParams.ParameterName, CommonParams.Value },
            Description = "Set parameter value"
        };

        public static readonly ActionDefinition Create = new()
        {
            Domain = Domain,
            Name = "create",
            Parameters = new[] { CommonParams.ExcelPath, CommonParams.ParameterName, CommonParams.Reference },
            Description = "Create new named range parameter"
        };

        public static readonly ActionDefinition Delete = new()
        {
            Domain = Domain,
            Name = "delete",
            Parameters = new[] { CommonParams.ExcelPath, CommonParams.ParameterName },
            Description = "Delete named range parameter"
        };

        public static IEnumerable<ActionDefinition> GetAll()
        {
            yield return List;
            yield return Get;
            yield return Set;
            yield return Create;
            yield return Delete;
        }

        public static ActionDefinition? GetByName(string actionName)
        {
            return GetAll().FirstOrDefault(a =>
                string.Equals(a.Name, actionName, StringComparison.OrdinalIgnoreCase));
        }
    }

    /// <summary>
    /// Table action definitions  
    /// </summary>
    public static class Table
    {
        private const string Domain = "Table";

        private static class CommonParams
        {
            public static readonly ParameterDefinition ExcelPath = SharedParams.ExcelPath;

            public static readonly ParameterDefinition TableName = new()
            {
                Name = "tableName",
                Required = true,
                MaxLength = 255,
                MinLength = 1,
                Description = "Excel table name"
            };

            public static readonly ParameterDefinition WorksheetName = new()
            {
                Name = "worksheetName",
                Required = true,
                MaxLength = 31,
                MinLength = 1,
                Pattern = @"^[^[\]/*?\\:]+$",
                Description = "Worksheet name"
            };

            public static readonly ParameterDefinition RangeAddress = new()
            {
                Name = "rangeAddress",
                Required = true,
                Pattern = @"^[A-Z]+[0-9]+:[A-Z]+[0-9]+$",
                Description = "Cell range address (e.g., A1:D10)"
            };

            public static readonly ParameterDefinition NewName = new()
            {
                Name = "newName",
                Required = true,
                MaxLength = 255,
                MinLength = 1,
                Description = "New table name"
            };
        }

        public static readonly ActionDefinition List = new()
        {
            Domain = Domain,
            Name = "list",
            Parameters = new[] { CommonParams.ExcelPath },
            Description = "List all tables in workbook"
        };

        public static readonly ActionDefinition Create = new()
        {
            Domain = Domain,
            Name = "create",
            Parameters = new[] 
            { 
                CommonParams.ExcelPath, 
                CommonParams.TableName,
                CommonParams.WorksheetName,
                CommonParams.RangeAddress
            },
            Description = "Create new Excel table from range"
        };

        public static readonly ActionDefinition Info = new()
        {
            Domain = Domain,
            Name = "info",
            Parameters = new[] { CommonParams.ExcelPath, CommonParams.TableName },
            Description = "Get table information (columns, row count, location)"
        };

        public static readonly ActionDefinition Rename = new()
        {
            Domain = Domain,
            Name = "rename",
            Parameters = new[] { CommonParams.ExcelPath, CommonParams.TableName, CommonParams.NewName },
            Description = "Rename table"
        };

        public static readonly ActionDefinition Delete = new()
        {
            Domain = Domain,
            Name = "delete",
            Parameters = new[] { CommonParams.ExcelPath, CommonParams.TableName },
            Description = "Delete table (keeps data, removes table structure)"
        };

        public static IEnumerable<ActionDefinition> GetAll()
        {
            yield return List;
            yield return Create;
            yield return Info;
            yield return Rename;
            yield return Delete;
        }

        public static ActionDefinition? GetByName(string actionName)
        {
            return GetAll().FirstOrDefault(a =>
                string.Equals(a.Name, actionName, StringComparison.OrdinalIgnoreCase));
        }
    }

    /// <summary>
    /// Lookup action definition by domain and name
    /// </summary>
    public static ActionDefinition? FindAction(string domain, string actionName)
    {
        return domain.ToLowerInvariant() switch
        {
            "powerquery" => PowerQuery.GetByName(actionName),
            "parameter" => Parameter.GetByName(actionName),
            "table" => Table.GetByName(actionName),
            _ => null
        };
    }

    /// <summary>
    /// Get all action definitions across all domains
    /// </summary>
    public static IEnumerable<ActionDefinition> GetAllActions()
    {
        foreach (var action in PowerQuery.GetAll()) yield return action;
        foreach (var action in Parameter.GetAll()) yield return action;
        foreach (var action in Table.GetAll()) yield return action;
    }
}
