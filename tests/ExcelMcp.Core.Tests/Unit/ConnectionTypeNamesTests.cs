using Sbroenne.ExcelMcp.Core.Connections;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connection")]
[Trait("RequiresExcel", "false")]
public sealed class ConnectionTypeNamesTests
{
    [Theory]
    [InlineData(Excel.XlConnectionType.xlConnectionTypeXMLMAP, "XMLMAP")]
    [InlineData(Excel.XlConnectionType.xlConnectionTypeTEXT, "TEXT")]
    [InlineData(Excel.XlConnectionType.xlConnectionTypeWEB, "WEB")]
    public void GetConnectionTypeName_MatchesExcelDefinitions(Excel.XlConnectionType type, string expected)
    {
        Assert.Equal(expected, ConnectionHelpers.GetConnectionTypeName((int)type));
    }
}
