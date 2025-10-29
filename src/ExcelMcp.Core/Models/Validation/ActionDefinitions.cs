using System.Text.RegularExpressions;

#pragma warning disable CS1591 // Missing XML comment - Description property documents each action

namespace Sbroenne.ExcelMcp.Core.Models.Validation;

/// <summary>
/// Central registry of all action definitions
/// Single source of truth for action metadata across CLI, MCP, and Core layers
/// </summary>
public static class ActionDefinitions
{
    /// <summary>
    /// Power Query action definitions
    /// </summary>
    public static class PowerQuery
    {
        private const string Tool = "excel_powerquery";
        private const string Domain = "PowerQuery";

        /// <summary>
        /// Common parameters used across multiple actions
        /// </summary>
        private static class CommonParams
        {
            public static readonly ParameterDefinition ExcelPath = new()
            {
                Name = "excelPath",
                Required = true,
                FileExtensions = new[] { "xlsx", "xlsm" },
                Description = "Excel file path (.xlsx or .xlsm)"
            };

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
            Action = "list",
            CliCommand = "pq-list",
            McpAction = "list",
            McpTool = Tool,
            Parameters = new[] { CommonParams.ExcelPath },
            Description = "List all Power Queries in workbook"
        };

        public static readonly ActionDefinition View = new()
        {
            Domain = Domain,
            Action = "view",
            CliCommand = "pq-view",
            McpAction = "view",
            McpTool = Tool,
            Parameters = new[] { CommonParams.ExcelPath, CommonParams.QueryName },
            Description = "View Power Query M code"
        };

        public static readonly ActionDefinition Import = new()
        {
            Domain = Domain,
            Action = "import",
            CliCommand = "pq-import",
            McpAction = "import",
            McpTool = Tool,
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
            Action = "export",
            CliCommand = "pq-export",
            McpAction = "export",
            McpTool = Tool,
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
            Action = "update",
            CliCommand = "pq-update",
            McpAction = "update",
            McpTool = Tool,
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
            Action = "delete",
            CliCommand = "pq-delete",
            McpAction = "delete",
            McpTool = Tool,
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
            Action = "refresh",
            CliCommand = "pq-refresh",
            McpAction = "refresh",
            McpTool = Tool,
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
            Action = "set-load-to-table",
            CliCommand = "pq-set-load-to-table",
            McpAction = "set-load-to-table",
            McpTool = Tool,
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
            Action = "set-load-to-data-model",
            CliCommand = "pq-set-load-to-data-model",
            McpAction = "set-load-to-data-model",
            McpTool = Tool,
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
            Action = "set-load-to-both",
            CliCommand = "pq-set-load-to-both",
            McpAction = "set-load-to-both",
            McpTool = Tool,
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
            Action = "set-connection-only",
            CliCommand = "pq-set-connection-only",
            McpAction = "set-connection-only",
            McpTool = Tool,
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
            Action = "get-load-config",
            CliCommand = "pq-get-load-config",
            McpAction = "get-load-config",
            McpTool = Tool,
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
        /// Gets action definition by MCP action name
        /// </summary>
        public static ActionDefinition? GetByMcpAction(string action)
        {
            return GetAll().FirstOrDefault(a => 
                string.Equals(a.McpAction, action, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Gets action definition by CLI command name
        /// </summary>
        public static ActionDefinition? GetByCliCommand(string command)
        {
            return GetAll().FirstOrDefault(a =>
                string.Equals(a.CliCommand, command, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Gets regex pattern for all valid MCP actions
        /// </summary>
        public static string GetMcpActionRegex()
        {
            var actions = GetAll().Select(a => Regex.Escape(a.McpAction));
            return $"^({string.Join("|", actions)})$";
        }
    }

    // TODO: Add Parameter, Table, DataModel, VBA, etc. action definitions
}
