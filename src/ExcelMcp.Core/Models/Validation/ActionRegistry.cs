namespace Sbroenne.ExcelMcp.Core.Models.Validation;

/// <summary>
/// Central registry of all action schemas
/// Single source of truth for actions across domains
/// </summary>
public static class ActionRegistry
{
    private static readonly ActionParameter ExcelPath = new()
    {
        Name = "excelPath",
        Required = true,
        Type = "string",
        Description = "Path to Excel file (.xlsx or .xlsm)"
    };
    
    private static readonly ActionParameter QueryName = new()
    {
        Name = "queryName",
        Required = true,
        Type = "string",
        MaxLength = 255,
        Description = "Name of Power Query"
    };
    
    private static readonly ActionParameter ParameterName = new()
    {
        Name = "parameterName",
        Required = true,
        Type = "string",
        MaxLength = 255,
        Description = "Name of parameter (named range)"
    };
    
    private static readonly ActionParameter TableName = new()
    {
        Name = "tableName",
        Required = true,
        Type = "string",
        MaxLength = 255,
        Description = "Name of Excel table"
    };
    
    public static class PowerQuery
    {
        public static readonly ActionSchema List = new()
        {
            Domain = "PowerQuery",
            Name = "list",
            Parameters = new[] { ExcelPath },
            Description = "List all Power Queries in workbook"
        };
        
        public static readonly ActionSchema View = new()
        {
            Domain = "PowerQuery",
            Name = "view",
            Parameters = new[] { ExcelPath, QueryName },
            Description = "View Power Query M code"
        };
        
        public static readonly ActionSchema Import = new()
        {
            Domain = "PowerQuery",
            Name = "import",
            Parameters = new[] { ExcelPath, QueryName, new ActionParameter { Name = "mCodeFile", Required = true } },
            Description = "Import Power Query from .pq file"
        };
        
        public static readonly ActionSchema Export = new()
        {
            Domain = "PowerQuery",
            Name = "export",
            Parameters = new[] { ExcelPath, QueryName, new ActionParameter { Name = "outputFile", Required = true } },
            Description = "Export Power Query M code to .pq file"
        };
        
        public static readonly ActionSchema Update = new()
        {
            Domain = "PowerQuery",
            Name = "update",
            Parameters = new[] { ExcelPath, QueryName, new ActionParameter { Name = "mCodeFile", Required = true } },
            Description = "Update Power Query M code"
        };
        
        public static readonly ActionSchema Delete = new()
        {
            Domain = "PowerQuery",
            Name = "delete",
            Parameters = new[] { ExcelPath, QueryName },
            Description = "Delete Power Query"
        };
        
        public static readonly ActionSchema Refresh = new()
        {
            Domain = "PowerQuery",
            Name = "refresh",
            Parameters = new[] { ExcelPath, QueryName },
            Description = "Refresh Power Query data"
        };
        
        public static readonly ActionSchema SetLoadToTable = new()
        {
            Domain = "PowerQuery",
            Name = "setLoadToTable",
            Parameters = new[] { ExcelPath, QueryName },
            Description = "Set query to load to worksheet table"
        };
        
        public static readonly ActionSchema SetLoadToDataModel = new()
        {
            Domain = "PowerQuery",
            Name = "setLoadToDataModel",
            Parameters = new[] { ExcelPath, QueryName },
            Description = "Set query to load to data model only"
        };
        
        public static readonly ActionSchema SetLoadToBoth = new()
        {
            Domain = "PowerQuery",
            Name = "setLoadToBoth",
            Parameters = new[] { ExcelPath, QueryName },
            Description = "Set query to load to both table and data model"
        };
        
        public static readonly ActionSchema SetConnectionOnly = new()
        {
            Domain = "PowerQuery",
            Name = "setConnectionOnly",
            Parameters = new[] { ExcelPath, QueryName },
            Description = "Set query as connection only (no load)"
        };
        
        public static readonly ActionSchema GetLoadConfig = new()
        {
            Domain = "PowerQuery",
            Name = "getLoadConfig",
            Parameters = new[] { ExcelPath, QueryName },
            Description = "Get query load configuration"
        };
    }
    
    public static class Parameter
    {
        public static readonly ActionSchema List = new()
        {
            Domain = "Parameter",
            Name = "list",
            Parameters = new[] { ExcelPath },
            Description = "List all parameters (named ranges)"
        };
        
        public static readonly ActionSchema Get = new()
        {
            Domain = "Parameter",
            Name = "get",
            Parameters = new[] { ExcelPath, ParameterName },
            Description = "Get parameter value"
        };
        
        public static readonly ActionSchema Set = new()
        {
            Domain = "Parameter",
            Name = "set",
            Parameters = new[] { ExcelPath, ParameterName, new ActionParameter { Name = "value", Required = true } },
            Description = "Set parameter value"
        };
        
        public static readonly ActionSchema Create = new()
        {
            Domain = "Parameter",
            Name = "create",
            Parameters = new[] { ExcelPath, ParameterName, new ActionParameter { Name = "value", Required = true } },
            Description = "Create new parameter"
        };
        
        public static readonly ActionSchema Delete = new()
        {
            Domain = "Parameter",
            Name = "delete",
            Parameters = new[] { ExcelPath, ParameterName },
            Description = "Delete parameter"
        };
    }
    
    public static class Table
    {
        public static readonly ActionSchema List = new()
        {
            Domain = "Table",
            Name = "list",
            Parameters = new[] { ExcelPath },
            Description = "List all Excel tables"
        };
        
        public static readonly ActionSchema Create = new()
        {
            Domain = "Table",
            Name = "create",
            Parameters = new[] { ExcelPath, TableName },
            Description = "Create new Excel table"
        };
        
        public static readonly ActionSchema Info = new()
        {
            Domain = "Table",
            Name = "info",
            Parameters = new[] { ExcelPath, TableName },
            Description = "Get table information"
        };
        
        public static readonly ActionSchema Rename = new()
        {
            Domain = "Table",
            Name = "rename",
            Parameters = new[] { ExcelPath, TableName, new ActionParameter { Name = "newName", Required = true } },
            Description = "Rename Excel table"
        };
        
        public static readonly ActionSchema Delete = new()
        {
            Domain = "Table",
            Name = "delete",
            Parameters = new[] { ExcelPath, TableName },
            Description = "Delete Excel table"
        };
    }
    
    public static ActionSchema? FindAction(string domain, string actionName)
    {
        var allActions = GetAllActions();
        return allActions.FirstOrDefault(a => 
            a.Domain.Equals(domain, StringComparison.OrdinalIgnoreCase) && 
            a.Name.Equals(actionName, StringComparison.OrdinalIgnoreCase));
    }
    
    public static IEnumerable<ActionSchema> GetActionsByDomain(string domain)
    {
        return GetAllActions().Where(a => a.Domain.Equals(domain, StringComparison.OrdinalIgnoreCase));
    }
    
    public static IEnumerable<ActionSchema> GetAllActions()
    {
        return new[]
        {
            PowerQuery.List, PowerQuery.View, PowerQuery.Import, PowerQuery.Export,
            PowerQuery.Update, PowerQuery.Delete, PowerQuery.Refresh,
            PowerQuery.SetLoadToTable, PowerQuery.SetLoadToDataModel, PowerQuery.SetLoadToBoth,
            PowerQuery.SetConnectionOnly, PowerQuery.GetLoadConfig,
            Parameter.List, Parameter.Get, Parameter.Set, Parameter.Create, Parameter.Delete,
            Table.List, Table.Create, Table.Info, Table.Rename, Table.Delete
        };
    }
}
