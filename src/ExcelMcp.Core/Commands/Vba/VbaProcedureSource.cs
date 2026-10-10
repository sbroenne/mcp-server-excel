using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace Sbroenne.ExcelMcp.Core.Commands;

internal sealed record VbaProcedureDeclaration(string Name, string Kind);

internal static class VbaProcedureSource
{
    private static readonly Regex ProcedureHeader = new(
        @"^\s*(?:(?:Public|Private|Friend|Global|Static)\s+)*(Sub|Function|Property\s+(?:Get|Let|Set))\s+([\p{L}_][\p{L}\p{Nd}_]*)\b",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex ProcedureEnd = new(
        @"^\s*End\s+(Sub|Function|Property)\s*$",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    public static VbaProcedureDeclaration ParseSingleProcedure(string source)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(source);

        string[] physicalLines = source.Split(["\r\n", "\n", "\r"], StringSplitOptions.None);
        string? firstCodeLine = physicalLines.FirstOrDefault(line => !string.IsNullOrWhiteSpace(line));
        string? lastCodeLine = physicalLines.LastOrDefault(line => !string.IsNullOrWhiteSpace(line));
        if (firstCodeLine == null || !ProcedureHeader.IsMatch(RemoveComment(firstCodeLine)) ||
            lastCodeLine == null || !ProcedureEnd.IsMatch(RemoveComment(lastCodeLine)))
        {
            throw new ArgumentException("Replacement code must start with and end at one VBA procedure.", nameof(source));
        }

        var statements = GetLogicalStatements(source);
        if (statements.Count < 2)
        {
            throw new ArgumentException("Replacement code must contain one complete VBA procedure.", nameof(source));
        }

        Match header = ProcedureHeader.Match(statements[0]);
        if (!header.Success)
        {
            throw new ArgumentException("Replacement code must start with one VBA procedure declaration.", nameof(source));
        }

        string kind = NormalizeKind(header.Groups[1].Value);
        string name = header.Groups[2].Value;

        for (int i = 1; i < statements.Count - 1; i++)
        {
            if (ProcedureHeader.IsMatch(statements[i]) || ProcedureEnd.IsMatch(statements[i]))
            {
                throw new ArgumentException("Replacement code must contain exactly one complete VBA procedure.", nameof(source));
            }
        }

        Match end = ProcedureEnd.Match(statements[^1]);
        if (!end.Success || !EndMatchesKind(end.Groups[1].Value, kind))
        {
            throw new ArgumentException("Replacement code must end with the matching End Sub, End Function, or End Property statement.", nameof(source));
        }

        return new VbaProcedureDeclaration(name, kind);
    }

    public static VbaProcedureDeclaration ParseProcedureHeader(string sourceLine)
    {
        return TryParseProcedureHeader(sourceLine)
            ?? throw new ArgumentException("VBA procedure declaration could not be read.", nameof(sourceLine));
    }

    public static VbaProcedureDeclaration? TryParseProcedureHeader(string sourceLine)
    {
        Match header = ProcedureHeader.Match(RemoveComment(sourceLine));
        return header.Success
            ? new VbaProcedureDeclaration(header.Groups[2].Value, NormalizeKind(header.Groups[1].Value))
            : null;
    }

    public static string ComputeHash(string source)
    {
        ArgumentNullException.ThrowIfNull(source);
        string normalized = source.Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n');
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(normalized)));
    }

    public static int GetBodyLineCount(string bodySource, string kind)
    {
        string[] lines = bodySource.Split(["\r\n", "\n", "\r"], StringSplitOptions.None);
        for (int index = 0; index < lines.Length; index++)
        {
            Match end = ProcedureEnd.Match(RemoveComment(lines[index]));
            if (end.Success && EndMatchesKind(end.Groups[1].Value, kind))
            {
                return index + 1;
            }
        }

        throw new ArgumentException("The existing procedure has no supported matching End statement; no source was changed.", nameof(bodySource));
    }

    public static string NormalizeKind(string kind)
    {
        string normalized = string.Join(' ', kind.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        return normalized.Equals("Sub", StringComparison.OrdinalIgnoreCase) ? "Sub" :
            normalized.Equals("Function", StringComparison.OrdinalIgnoreCase) ? "Function" :
            normalized.Equals("Property Get", StringComparison.OrdinalIgnoreCase) ? "Property Get" :
            normalized.Equals("Property Let", StringComparison.OrdinalIgnoreCase) ? "Property Let" :
            normalized.Equals("Property Set", StringComparison.OrdinalIgnoreCase) ? "Property Set" :
            string.Empty;
    }

    private static bool EndMatchesKind(string endKind, string procedureKind)
    {
        if (procedureKind.StartsWith("Property ", StringComparison.Ordinal))
        {
            return endKind.Equals("Property", StringComparison.OrdinalIgnoreCase);
        }

        return endKind.Equals(procedureKind, StringComparison.OrdinalIgnoreCase);
    }

    private static List<string> GetLogicalStatements(string source)
    {
        var statements = new List<string>();
        var current = new StringBuilder();

        foreach (string physicalLine in source.Split(["\r\n", "\n", "\r"], StringSplitOptions.None))
        {
            string code = RemoveComment(physicalLine).Trim();
            if (code.Length == 0 || Regex.IsMatch(code, @"^Rem(?:\s|$)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                continue;
            }

            bool continues = HasLineContinuation(code);
            if (continues)
            {
                code = code[..^1].TrimEnd();
            }

            if (current.Length > 0)
            {
                current.Append(' ');
            }

            current.Append(code);
            if (!continues)
            {
                statements.Add(current.ToString());
                current.Clear();
            }
        }

        if (current.Length > 0)
        {
            throw new ArgumentException("Replacement code ends with an incomplete continued line.", nameof(source));
        }

        return statements;
    }

    private static string RemoveComment(string line)
    {
        bool inString = false;
        for (int i = 0; i < line.Length; i++)
        {
            if (line[i] == '"')
            {
                if (inString && i + 1 < line.Length && line[i + 1] == '"')
                {
                    i++;
                }
                else
                {
                    inString = !inString;
                }
            }
            else if (line[i] == '\'' && !inString)
            {
                return line[..i];
            }
        }

        return line;
    }

    private static bool HasLineContinuation(string line)
    {
        if (!line.EndsWith('_'))
        {
            return false;
        }

        bool inString = false;
        for (int i = 0; i < line.Length - 1; i++)
        {
            if (line[i] == '"')
            {
                if (inString && i + 1 < line.Length - 1 && line[i + 1] == '"')
                {
                    i++;
                }
                else
                {
                    inString = !inString;
                }
            }
        }

        return !inString && (line.Length == 1 || char.IsWhiteSpace(line[^2]));
    }
}
