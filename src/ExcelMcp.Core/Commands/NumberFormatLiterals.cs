using System.Text;

namespace Sbroenne.ExcelMcp.Core.Commands;

internal static class NumberFormatLiterals
{
    internal static string PreserveCurrencyLiterals(string format)
    {
        if (!format.Contains('$'))
        {
            return format;
        }

        var result = new StringBuilder(format.Length);
        var quoted = false;
        var bracketed = false;
        for (var index = 0; index < format.Length; index++)
        {
            var character = format[index];
            if (!quoted && !bracketed && character == '"' &&
                index + 2 < format.Length && format[index + 1] == '$' && format[index + 2] == '"')
            {
                result.Append("\\$");
                index += 2;
                continue;
            }
            if (!quoted && !bracketed && character is '\\' or '_' or '*')
            {
                result.Append(character);
                if (index + 1 < format.Length)
                {
                    result.Append(format[++index]);
                }
                continue;
            }
            if (!bracketed && character == '"')
            {
                quoted = !quoted;
            }
            else if (!quoted && character == '[')
            {
                bracketed = true;
            }
            else if (!quoted && character == ']')
            {
                bracketed = false;
            }
            else if (!quoted && !bracketed && character == '$')
            {
                // Invariant Excel format properties can substitute the regional currency.
                result.Append('\\');
            }
            result.Append(character);
        }
        return result.ToString();
    }
}
