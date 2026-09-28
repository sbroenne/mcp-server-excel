using Sbroenne.ExcelMcp.Core.Commands.XmlMap;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "XmlMap")]
[Trait("RequiresExcel", "true")]
public sealed partial class PersistentServiceXmlMapTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private const string CustomerSchema = """
        <?xml version="1.0" encoding="utf-8"?>
        <xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema">
          <xs:element name="customer">
            <xs:complexType>
              <xs:sequence>
                <xs:element name="name" type="xs:string" />
              </xs:sequence>
            </xs:complexType>
          </xs:element>
        </xs:schema>
        """;

    private readonly IXmlMapCommands _xmlMapCommands =
        ServiceCommandProxy.Create<IXmlMapCommands>(fixture);
}
