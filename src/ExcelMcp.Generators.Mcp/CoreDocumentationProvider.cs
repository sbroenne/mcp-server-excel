using System.Globalization;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Xml.Linq;
using Microsoft.CodeAnalysis;

namespace Sbroenne.ExcelMcp.Generators.Mcp;

internal sealed class CoreDocumentationProvider(string xml) : DocumentationProvider
{
    private readonly Dictionary<string, string> _members = XDocument.Parse(xml)
        .Descendants("member")
        .ToDictionary(member => (string)member.Attribute("name")!, member => member.ToString());

    protected override string GetDocumentationForSymbol(
        string documentationMemberID,
        CultureInfo preferredCulture,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return _members.TryGetValue(documentationMemberID, out var documentation) ? documentation : string.Empty;
    }

    public override bool Equals(object? obj) => ReferenceEquals(this, obj);

    public override int GetHashCode() => RuntimeHelpers.GetHashCode(this);
}
