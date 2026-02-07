using System.Text;
using Microsoft.CodeAnalysis;

namespace Sbroenne.ExcelMcp.Generators.Common;

/// <summary>
/// String manipulation utilities shared between generators.
/// </summary>
public static class StringHelper
{
    public static string ToKebabCase(string pascalCase)
    {
        var sb = new StringBuilder();
        for (int i = 0; i < pascalCase.Length; i++)
        {
            var c = pascalCase[i];
            if (char.IsUpper(c))
            {
                if (i > 0)
                    sb.Append('-');
                sb.Append(char.ToLowerInvariant(c));
            }
            else
            {
                sb.Append(c);
            }
        }
        return sb.ToString();
    }

    public static string ToPascalCase(string kebabCase)
    {
        var parts = kebabCase.Split('-');
        return string.Concat(parts.Select(p =>
            p.Length > 0 ? char.ToUpperInvariant(p[0]) + p.Substring(1) : p));
    }

    /// <summary>
    /// Converts camelCase or PascalCase to snake_case.
    /// Example: "sheetName" → "sheet_name", "rangeAddress" → "range_address"
    /// </summary>
    public static string ToSnakeCase(string camelCase)
    {
        var sb = new StringBuilder();
        for (int i = 0; i < camelCase.Length; i++)
        {
            var c = camelCase[i];
            if (char.IsUpper(c))
            {
                if (i > 0)
                    sb.Append('_');
                sb.Append(char.ToLowerInvariant(c));
            }
            else
            {
                sb.Append(c);
            }
        }
        return sb.ToString();
    }

    public static bool IsStringType(string typeName)
    {
        var normalized = typeName.TrimEnd('?');
        return normalized == "string" || normalized == "System.String";
    }

    /// <summary>
    /// Gets CLI option description based on parameter name.
    /// </summary>
    public static string GetParameterDescription(string paramName)
    {
        return paramName switch
        {
            "queryName" => "Query name",
            "mCode" => "M code formula",
            "mCodeFile" => "Path to file containing M code",
            "loadDestination" => "Load destination: worksheet, data-model, both, connection-only",
            "targetSheet" => "Target worksheet name",
            "targetCellAddress" => "Target cell address (e.g., A1)",
            "oldName" => "Current name (for rename)",
            "newName" => "New name (for rename)",
            "timeout" => "Timeout duration",
            "refresh" => "Whether to refresh after update",
            "sheetName" => "Worksheet name",
            "tableName" => "Table name",
            "connectionName" => "Connection name",
            "chartName" => "Chart name",
            "slicerName" => "Slicer name",
            "pivotTableName" => "PivotTable name",
            "rangeAddress" => "Range address (e.g., A1:C10)",
            "values" => "Values to set",
            "formula" => "Formula to apply",
            "format" => "Format string",
            _ => ToPascalCase(paramName)
        };
    }
}

/// <summary>
/// Type name handling utilities.
/// </summary>
public static class TypeNameHelper
{
    public static string GetTypeName(ITypeSymbol type, NullableAnnotation nullableAnnotation = NullableAnnotation.None)
    {
        if (type.SpecialType == SpecialType.System_Void)
            return "void";
        if (type.SpecialType == SpecialType.System_String)
        {
            return nullableAnnotation == NullableAnnotation.Annotated ? "string?" : "string";
        }
        if (type.SpecialType == SpecialType.System_Boolean)
            return "bool";
        if (type.SpecialType == SpecialType.System_Int32)
            return "int";

        if (type is INamedTypeSymbol namedType)
        {
            if (namedType.IsGenericType && namedType.ConstructedFrom.SpecialType == SpecialType.System_Nullable_T)
            {
                return GetTypeName(namedType.TypeArguments[0]) + "?";
            }
        }

        // For custom types, use fully qualified name
        var fullName = type.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
        // Remove "global::" prefix
        if (fullName.StartsWith("global::"))
            fullName = fullName.Substring(8);

        return fullName;
    }

    public static string GetDefaultValueString(IParameterSymbol param)
    {
        if (!param.HasExplicitDefaultValue)
            return "default";

        var value = param.ExplicitDefaultValue;

        if (value is null)
            return "null";
        if (value is bool b)
            return b ? "true" : "false";
        if (value is string s)
            return $"\"{s}\"";
        if (value is int or long or short or byte)
        {
            // Handle enum defaults
            if (param.Type.TypeKind == TypeKind.Enum)
            {
                var enumType = param.Type.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
                if (enumType.StartsWith("global::"))
                    enumType = enumType.Substring(8);

                // Find the enum member with this value
                foreach (var member in param.Type.GetMembers())
                {
                    if (member is IFieldSymbol field && field.HasConstantValue &&
                        field.ConstantValue?.Equals(value) == true)
                    {
                        return $"{enumType}.{field.Name}";
                    }
                }
                // Fallback to numeric value
                return $"({enumType}){value}";
            }
            return value.ToString()!;
        }

        // Handle double/float defaults — ensure decimal point so C# treats as double literal
        // Without this, double width = 400 → DefaultValue(400) → int → InvalidCastException
        if (value is double d)
        {
            var ds = d.ToString("G", System.Globalization.CultureInfo.InvariantCulture);
            if (!ds.Contains('.') && !ds.Contains('E') && !ds.Contains('e'))
                ds += ".0";
            return ds;
        }
        if (value is float f)
        {
            var fs = f.ToString("G", System.Globalization.CultureInfo.InvariantCulture);
            if (!fs.Contains('.') && !fs.Contains('E') && !fs.Contains('e'))
                fs += ".0";
            return fs + "f";
        }

        // Handle enum defaults
        if (param.Type.TypeKind == TypeKind.Enum)
        {
            return $"{param.Type.Name}.{value}";
        }

        return value.ToString() ?? "default";
    }
}
