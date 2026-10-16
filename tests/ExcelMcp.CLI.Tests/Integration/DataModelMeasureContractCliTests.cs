using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// CLI arguments for DAX measure writes. Measure creation, update, and
/// evaluation in Excel are covered by the Service Data Model tests.
/// </summary>
[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "DataModel")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class DataModelMeasureContractCliTests
{
    private const string Formula = "DIVIDE(SUM(SalesTable[Amount]), 1000)";

    [Fact]
    public async Task CreateAndUpdateMeasure_SendFormulaUnchanged()
    {
        var requests = new List<ServiceRequest>();
        ServiceResponse Respond(ServiceRequest request)
        {
            requests.Add(request);
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        }

        var created = await InProcessCliHelper.RunAsync(
            ["datamodel", "create-measure", "--session", "session-1", "--table-name", "SalesTable",
             "--measure-name", "Total", "--dax-formula", Formula, "--format-type", "Decimal"], Respond);
        var updated = await InProcessCliHelper.RunAsync(
            ["datamodel", "update-measure", "--session", "session-1", "--measure-name", "Total",
             "--dax-formula", Formula, "--format-type", "Decimal"], Respond);

        Assert.Equal(0, created.ExitCode);
        Assert.Equal(0, updated.ExitCode);
        Assert.Collection(requests,
            request => AssertRequest(request, "datamodel.create-measure",
                $$"""{"tableName":"SalesTable","measureName":"Total","daxFormula":{{JsonSerializer.Serialize(Formula)}},"formatType":"Decimal"}"""),
            request => AssertRequest(request, "datamodel.update-measure",
                $$"""{"measureName":"Total","daxFormula":{{JsonSerializer.Serialize(Formula)}},"formatType":"Decimal"}"""));
    }

    private static void AssertRequest(ServiceRequest request, string command, string expectedArgs)
    {
        Assert.Equal(command, request.Command);
        Assert.Equal("session-1", request.SessionId);
        Assert.Equal(JsonDocument.Parse(expectedArgs).RootElement.ToString(),
            JsonDocument.Parse(request.Args!).RootElement.ToString());
    }
}
