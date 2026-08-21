using System.Reflection;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Analysis;
using Sbroenne.ExcelMcp.Core.Commands.Calculation;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.XmlMap;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Sbroenne.ExcelMcp.Generated;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "GeneratedContracts")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class GeneratedActionContractTests
{
    [Theory]
    [InlineData("worksheet", PowerQueryLoadMode.LoadToTable)]
    [InlineData("TABLE", PowerQueryLoadMode.LoadToTable)]
    [InlineData("data-model", PowerQueryLoadMode.LoadToDataModel)]
    [InlineData("both", PowerQueryLoadMode.LoadToBoth)]
    [InlineData("connection-only", PowerQueryLoadMode.ConnectionOnly)]
    [InlineData("load-to-table", PowerQueryLoadMode.LoadToTable)]
    public void PowerQueryDispatch_ParsesDocumentedLoadDestinationAliases(
        string suppliedValue,
        PowerQueryLoadMode expected)
    {
        var (commands, proxy) = CreateProxy<IPowerQueryCommands>();

        ServiceRegistry.PowerQuery.DispatchToCore(
            commands,
            PowerQueryAction.LoadTo,
            null!,
            JsonSerializer.Serialize(new
            {
                queryName = "Probe",
                loadDestination = suppliedValue
            }));

        Assert.Equal(1, proxy.CallCount);
        Assert.Equal(expected, Assert.IsType<PowerQueryLoadMode>(proxy.LastArguments![2]));
    }

    [Theory]
    [InlineData("not-a-destination")]
    [InlineData("work_sheet")]
    [InlineData("work-sheet")]
    [InlineData("0")]
    [InlineData("999")]
    [InlineData("-1")]
    public void PowerQueryDispatch_RejectsUnknownLoadDestinationBeforeCoreDispatch(string suppliedValue)
    {
        var (commands, proxy) = CreateProxy<IPowerQueryCommands>();

        var exception = Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.PowerQuery.DispatchToCore(
                commands,
                PowerQueryAction.LoadTo,
                null!,
                JsonSerializer.Serialize(new { queryName = "Probe", loadDestination = suppliedValue })));

        Assert.Contains("loadDestination", exception.Message, StringComparison.Ordinal);
        Assert.Equal(0, proxy.CallCount);
    }

    [Fact]
    public void LegacyPowerQueryLoadModeParser_RejectsUnknownValue()
    {
        Assert.Equal(PowerQueryLoadMode.LoadToTable, ParameterTransforms.ParseLoadMode("worksheet"));
        Assert.Throws<ArgumentException>(() => ParameterTransforms.ParseLoadMode("not-a-destination"));
    }

    [Theory]
    [InlineData("set-mode", """{"mode":"not-a-mode"}""")]
    [InlineData("calculate", """{"scope":"not-a-scope"}""")]
    public void CalculationDispatch_RejectsUnknownEnumsBeforeCoreDispatch(string action, string argsJson)
    {
        var (commands, proxy) = CreateProxy<ICalculationModeCommands>();
        Assert.True(ServiceRegistry.Calculation.TryParseAction(action, out var parsedAction));

        Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.Calculation.DispatchToCore(commands, parsedAction, null!, argsJson));

        Assert.Equal(0, proxy.CallCount);
    }

    [Fact]
    public void AnalysisDispatch_RejectsUnknownScenarioSummaryTypeBeforeCoreDispatch()
    {
        var (commands, proxy) = CreateProxy<IAnalysisCommands>();

        Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.Analysis.DispatchToCore(
                commands,
                AnalysisAction.CreateScenarioSummary,
                null!,
                """{"sheetName":"Model","reportType":"not-a-report"}"""));

        Assert.Equal(0, proxy.CallCount);
    }

    [Fact]
    public void ChartDispatch_RejectsUnknownChartTypeBeforeCoreDispatch()
    {
        var (commands, proxy) = CreateProxy<IChartCommands>();

        Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.Chart.DispatchToCore(
                commands,
                ChartAction.CreateFromRange,
                null!,
                """{"sheetName":"Model","sourceRangeAddress":"A1:B2","chartType":"not-a-chart"}"""));

        Assert.Equal(0, proxy.CallCount);
    }

    [Fact]
    public void PivotTableDispatch_RejectsUnknownNullableEnumBeforeCoreDispatch()
    {
        var (commands, proxy) = CreateProxy<IPivotTableCommands>();

        Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.PivotTable.DispatchToCore(
                commands,
                PivotTableAction.SetCacheOptions,
                null!,
                """{"pivotTableName":"PivotTable1","missingItemsLimit":"not-a-limit"}"""));

        Assert.Equal(0, proxy.CallCount);
    }

    [Theory]
    [InlineData("calculate", "mode", "manual")]
    [InlineData("get-mode", "mode", "manual")]
    public void CalculationCliRoute_RejectsParametersFromOtherActions(
        string action,
        string parameterName,
        string parameterValue)
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.Calculation.RouteCliArgs(
                action,
                mode: parameterValue,
                scope: action == "calculate" ? "workbook" : null));

        Assert.Contains(parameterName, exception.Message, StringComparison.Ordinal);
        Assert.Contains(action, exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerQueryCliRoute_RejectsParametersFromOtherActions()
    {
        var deleteException = Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.PowerQuery.RouteCliArgs(
                "delete",
                queryName: "Probe",
                mCode: "let Source = 1 in Source"));
        Assert.Contains("mCode", deleteException.Message, StringComparison.Ordinal);

        var loadToException = Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.PowerQuery.RouteCliArgs(
                "load-to",
                queryName: "Probe",
                loadDestination: "worksheet",
                timeout: 30));
        Assert.Contains("timeout", loadToException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerQueryCliRoute_DoesNotTreatOmittedActionDefaultsAsSupplied()
    {
        var route = ServiceRegistry.PowerQuery.RouteCliArgs("delete", queryName: "Probe");

        Assert.Equal("powerquery.delete", route.Command);
    }

    [Theory]
    [InlineData("worksheet")]
    [InlineData("WORKSHEET")]
    [InlineData("data-model")]
    [InlineData("DATA-MODEL")]
    public void PowerQueryCliRoute_AcceptsExactAliasesIgnoringCase(string suppliedValue)
    {
        var route = ServiceRegistry.PowerQuery.RouteCliArgs(
            "load-to",
            queryName: "Probe",
            loadDestination: suppliedValue);

        Assert.Equal("powerquery.load-to", route.Command);
    }

    [Fact]
    public void RawActionValidation_RejectsActionInapplicableExplicitNull()
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.ValidateCommandArguments(
                "calculation.calculate",
                """{"scope":"workbook","mode":null}"""));

        Assert.Contains("mode", exception.Message, StringComparison.Ordinal);
        Assert.Contains("calculate", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("""{"sheetName":null,"rangeAddress":"A1"}""")]
    [InlineData("""{"sheetName":42,"rangeAddress":"A1"}""")]
    [InlineData("""{"sheetName":true,"rangeAddress":"A1"}""")]
    public void AllowEmptyRequiredString_RejectsNonStringJsonBeforeCoreDispatch(
        string argsJson)
    {
        var (commands, proxy) = CreateProxy<IConditionalFormattingCommands>();

        var exception = Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.ConditionalFormat.DispatchToCore(
                commands,
                ConditionalFormatAction.ClearRules,
                null!,
                argsJson));

        Assert.Contains("sheetName", exception.Message, StringComparison.Ordinal);
        Assert.Contains("JSON string", exception.Message, StringComparison.Ordinal);
        Assert.Equal(0, proxy.CallCount);
    }

    [Fact]
    public void AllowEmptyRequiredString_AcceptsEmptyJsonString()
    {
        var (commands, proxy) = CreateProxy<IConditionalFormattingCommands>();

        ServiceRegistry.ConditionalFormat.DispatchToCore(
            commands,
            ConditionalFormatAction.ClearRules,
            null!,
            """{"sheetName":"","rangeAddress":"A1"}""");

        Assert.Equal(1, proxy.CallCount);
        Assert.Equal(string.Empty, proxy.LastArguments![1]);
    }

    [Theory]
    [InlineData(
        "powerquery.refresh",
        """{"queryName":"Probe","Timeout":60}""",
        "Timeout")]
    [InlineData(
        "powerquery.evaluate",
        """{"MCodeFile":"query.m"}""",
        "MCodeFile")]
    [InlineData(
        "vba.import",
        """{"moduleName":"Module1","VbaCodeFile":"module.bas"}""",
        "VbaCodeFile")]
    [InlineData(
        "datamodel.create-measure",
        """{"tableName":"Sales","measureName":"Total","DaxFormulaFile":"measure.dax"}""",
        "DaxFormulaFile")]
    [InlineData(
        "datamodel.evaluate",
        """{"DaxQueryFile":"query.dax"}""",
        "DaxQueryFile")]
    [InlineData(
        "datamodel.execute-dmv",
        """{"DmvQueryFile":"query.dmv"}""",
        "DmvQueryFile")]
    [InlineData(
        "xmlmap.add",
        """{"SchemaFile":"schema.xsd"}""",
        "SchemaFile")]
    [InlineData(
        "xmlmap.import-xml",
        """{"XmlDataFile":"data.xml"}""",
        "XmlDataFile")]
    public void RawActionValidation_RejectsNonCanonicalPropertyCasing(
        string command,
        string argsJson,
        string suppliedProperty)
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.ValidateCommandArguments(command, argsJson));

        Assert.Contains(suppliedProperty, exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("""{"queryName":"Probe","Timeout":null}""", "Timeout")]
    [InlineData("""{"queryName":"Probe","unexpected":null}""", "unexpected")]
    public void RawActionValidation_RejectsInvalidNullPropertyNames(
        string argsJson,
        string suppliedProperty)
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.ValidateCommandArguments("powerquery.refresh", argsJson));

        Assert.Contains(suppliedProperty, exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void FileOrValueDispatch_ResolvesEveryGeneratedFileAlias()
    {
        var path = Path.GetTempFileName();
        const string content = "canonical file content";
        File.WriteAllText(path, content);

        try
        {
            AssertResolvedFileArgument<IPowerQueryCommands, PowerQueryAction>(
                ServiceRegistry.PowerQuery.DispatchToCore,
                PowerQueryAction.Evaluate,
                JsonSerializer.Serialize(new { mCodeFile = path }),
                expectedArgumentIndex: 1,
                content);
            AssertResolvedFileArgument<IVbaCommands, VbaAction>(
                ServiceRegistry.Vba.DispatchToCore,
                VbaAction.Import,
                JsonSerializer.Serialize(new { moduleName = "Module1", vbaCodeFile = path }),
                expectedArgumentIndex: 2,
                content);
            AssertResolvedFileArgument<IDataModelCommands, DataModelAction>(
                ServiceRegistry.DataModel.DispatchToCore,
                DataModelAction.CreateMeasure,
                JsonSerializer.Serialize(new
                {
                    tableName = "Sales",
                    measureName = "Total",
                    daxFormulaFile = path
                }),
                expectedArgumentIndex: 3,
                content);
            AssertResolvedFileArgument<IDataModelCommands, DataModelAction>(
                ServiceRegistry.DataModel.DispatchToCore,
                DataModelAction.Evaluate,
                JsonSerializer.Serialize(new { daxQueryFile = path }),
                expectedArgumentIndex: 1,
                content);
            AssertResolvedFileArgument<IDataModelCommands, DataModelAction>(
                ServiceRegistry.DataModel.DispatchToCore,
                DataModelAction.ExecuteDmv,
                JsonSerializer.Serialize(new { dmvQueryFile = path }),
                expectedArgumentIndex: 1,
                content);
            AssertResolvedFileArgument<IXmlMapCommands, XmlMapAction>(
                ServiceRegistry.XmlMap.DispatchToCore,
                XmlMapAction.Add,
                JsonSerializer.Serialize(new { schemaFile = path }),
                expectedArgumentIndex: 1,
                content);
            AssertResolvedFileArgument<IXmlMapCommands, XmlMapAction>(
                ServiceRegistry.XmlMap.DispatchToCore,
                XmlMapAction.ImportXml,
                JsonSerializer.Serialize(new { xmlDataFile = path }),
                expectedArgumentIndex: 1,
                content);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void FileOrValueDispatch_RejectsConflictingInlineAndFileInputs()
    {
        var path = Path.GetTempFileName();
        try
        {
            var (commands, proxy) = CreateProxy<IPowerQueryCommands>();

            var exception = Assert.Throws<ArgumentException>(() =>
                ServiceRegistry.PowerQuery.DispatchToCore(
                    commands,
                    PowerQueryAction.Evaluate,
                    null!,
                    JsonSerializer.Serialize(new
                    {
                        mCode = "let Source = 1 in Source",
                        mCodeFile = path
                    })));

            Assert.Contains("mCode", exception.Message, StringComparison.Ordinal);
            Assert.Contains("mCodeFile", exception.Message, StringComparison.Ordinal);
            Assert.Equal(0, proxy.CallCount);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void FileOrValueDispatch_RejectsMissingRequiredInput()
    {
        var (commands, proxy) = CreateProxy<IPowerQueryCommands>();

        var exception = Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.PowerQuery.DispatchToCore(
                commands,
                PowerQueryAction.Evaluate,
                null!,
                "{}"));

        Assert.Contains("mCode", exception.Message, StringComparison.Ordinal);
        Assert.Equal(0, proxy.CallCount);
    }

    [Fact]
    public void FileOrValueDispatch_AllowsNeitherForOptionalInput()
    {
        var (commands, proxy) = CreateProxy<IDataModelCommands>();

        ServiceRegistry.DataModel.DispatchToCore(
            commands,
            DataModelAction.UpdateMeasure,
            null!,
            """{"measureName":"Total","description":"Updated"}""");

        Assert.Equal(1, proxy.CallCount);
        Assert.Null(proxy.LastArguments![2]);
    }

    [Fact]
    public void FileOrValueDispatch_RejectsMissingAndUnreadableFiles()
    {
        var missingPath = Path.Join(Path.GetTempPath(), $"{Guid.NewGuid():N}.m");
        var (missingCommands, missingProxy) = CreateProxy<IPowerQueryCommands>();
        Assert.Throws<FileNotFoundException>(() =>
            ServiceRegistry.PowerQuery.DispatchToCore(
                missingCommands,
                PowerQueryAction.Evaluate,
                null!,
                JsonSerializer.Serialize(new { mCodeFile = missingPath })));
        Assert.Equal(0, missingProxy.CallCount);

        var unreadablePath = Path.GetTempFileName();
        try
        {
            using var lockStream = new FileStream(
                unreadablePath,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.None);
            var (unreadableCommands, unreadableProxy) = CreateProxy<IPowerQueryCommands>();
            Assert.Throws<IOException>(() =>
                ServiceRegistry.PowerQuery.DispatchToCore(
                    unreadableCommands,
                    PowerQueryAction.Evaluate,
                    null!,
                    JsonSerializer.Serialize(new { mCodeFile = unreadablePath })));
            Assert.Equal(0, unreadableProxy.CallCount);
        }
        finally
        {
            File.Delete(unreadablePath);
        }
    }

    [Fact]
    public void RawDispatch_ConvertsIntegerTimeoutSecondsExactlyOnce()
    {
        var (powerQueryCommands, powerQueryProxy) = CreateProxy<IPowerQueryCommands>();
        ServiceRegistry.PowerQuery.DispatchToCore(
            powerQueryCommands,
            PowerQueryAction.Refresh,
            null!,
            """{"queryName":"Probe","timeout":0}""");
        Assert.Equal(TimeSpan.Zero, Assert.IsType<TimeSpan>(powerQueryProxy.LastArguments![2]));

        var (connectionCommands, connectionProxy) = CreateProxy<IConnectionCommands>();
        ServiceRegistry.Connection.DispatchToCore(
            connectionCommands,
            ConnectionAction.Refresh,
            null!,
            """{"connectionName":"Probe","timeout":1}""");
        Assert.Equal(TimeSpan.FromSeconds(1), Assert.IsType<TimeSpan>(connectionProxy.LastArguments![2]));

        var (dataModelCommands, dataModelProxy) = CreateProxy<IDataModelCommands>();
        ServiceRegistry.DataModel.DispatchToCore(
            dataModelCommands,
            DataModelAction.Refresh,
            null!,
            """{"timeout":2147483}""");
        Assert.Equal(TimeSpan.FromSeconds(2147483), Assert.IsType<TimeSpan>(dataModelProxy.LastArguments![2]));

        var (pivotCommands, pivotProxy) = CreateProxy<IPivotTableCommands>();
        ServiceRegistry.PivotTable.DispatchToCore(
            pivotCommands,
            PivotTableAction.Refresh,
            null!,
            """{"pivotTableName":"Probe","timeout":60}""");
        Assert.Equal(TimeSpan.FromSeconds(60), Assert.IsType<TimeSpan>(pivotProxy.LastArguments![2]));

        var (vbaCommands, vbaProxy) = CreateProxy<IVbaCommands>();
        ServiceRegistry.Vba.DispatchToCore(
            vbaCommands,
            VbaAction.Run,
            null!,
            """{"procedureName":"Probe","timeout":2147483}""");
        Assert.Equal(TimeSpan.FromSeconds(2147483), Assert.IsType<TimeSpan>(vbaProxy.LastArguments![2]));
    }

    [Theory]
    [InlineData("""{"queryName":"Probe","timeout":-1}""")]
    [InlineData("""{"queryName":"Probe","timeout":2147484}""")]
    [InlineData("""{"queryName":"Probe","timeout":"600"}""")]
    public void RawDispatch_RejectsInvalidTimeoutSecondsBeforeCoreDispatch(string argsJson)
    {
        var (commands, proxy) = CreateProxy<IPowerQueryCommands>();

        Assert.ThrowsAny<Exception>(() =>
            ServiceRegistry.PowerQuery.DispatchToCore(
                commands,
                PowerQueryAction.Refresh,
                null!,
                argsJson));

        Assert.Equal(0, proxy.CallCount);
    }

    [Fact]
    public void RawDispatch_PreservesMissingAndZeroTimeoutSemantics()
    {
        var (powerQueryCommands, powerQueryProxy) = CreateProxy<IPowerQueryCommands>();
        ServiceRegistry.PowerQuery.DispatchToCore(
            powerQueryCommands,
            PowerQueryAction.Refresh,
            null!,
            """{"queryName":"Probe"}""");
        Assert.Equal(TimeSpan.Zero, Assert.IsType<TimeSpan>(powerQueryProxy.LastArguments![2]));

        var (connectionCommands, connectionProxy) = CreateProxy<IConnectionCommands>();
        ServiceRegistry.Connection.DispatchToCore(
            connectionCommands,
            ConnectionAction.Refresh,
            null!,
            """{"connectionName":"Probe"}""");
        Assert.Null(connectionProxy.LastArguments![2]);
    }

    [Fact]
    public void RawDispatch_RejectsZeroForNonPowerQueryTimeouts()
    {
        var (connectionCommands, connectionProxy) = CreateProxy<IConnectionCommands>();
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            ServiceRegistry.Connection.DispatchToCore(
                connectionCommands,
                ConnectionAction.Refresh,
                null!,
                """{"connectionName":"Probe","timeout":0}"""));
        Assert.Equal(0, connectionProxy.CallCount);

        var (dataModelCommands, dataModelProxy) = CreateProxy<IDataModelCommands>();
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            ServiceRegistry.DataModel.DispatchToCore(
                dataModelCommands,
                DataModelAction.Refresh,
                null!,
                """{"timeout":0}"""));
        Assert.Equal(0, dataModelProxy.CallCount);
    }

    [Fact]
    public void GeneratedRoutes_RejectOutOfRangeTimeoutsBeforeForwarding()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            ServiceRegistry.PowerQuery.RouteCliArgs(
                "refresh",
                queryName: "Probe",
                timeout: -1));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            ServiceRegistry.Connection.RouteCliArgs(
                "refresh",
                connectionName: "Probe",
                timeout: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            ServiceRegistry.Vba.RouteCliArgs(
                "run",
                procedureName: "Probe",
                timeout: ParameterTransforms.MaximumTimeoutSeconds + 1));
    }

    [Theory]
    [InlineData(10)]
    [InlineData(3600)]
    public void SessionTimeoutValidation_AcceptsDocumentedBoundaries(int timeoutSeconds)
    {
        Assert.Equal(
            TimeSpan.FromSeconds(timeoutSeconds),
            ParameterTransforms.ParseTimeoutSeconds(
                timeoutSeconds,
                "timeoutSeconds",
                minimumSeconds: 10,
                maximumSeconds: 3600));
    }

    [Fact]
    public void GeneratedPublicTimeoutContract_UsesNullableIntegerSeconds()
    {
        var routeParameter = typeof(ServiceRegistry.PowerQuery)
            .GetMethod(nameof(ServiceRegistry.PowerQuery.RouteCliArgs))!
            .GetParameters()
            .Single(parameter => parameter.Name == "timeout");
        var actionParameter = typeof(ServiceRegistry.PowerQuery)
            .GetMethod(nameof(ServiceRegistry.PowerQuery.RouteAction))!
            .GetParameters()
            .Single(parameter => parameter.Name == "timeout");

        Assert.Equal(typeof(int?), routeParameter.ParameterType);
        Assert.Equal(typeof(int?), actionParameter.ParameterType);
    }

    private static void AssertResolvedFileArgument<TInterface, TAction>(
        Func<TInterface, TAction, IExcelBatch, string?, string?> dispatch,
        TAction action,
        string argsJson,
        int expectedArgumentIndex,
        string expectedContent)
        where TInterface : class
        where TAction : struct, Enum
    {
        var (commands, proxy) = CreateProxy<TInterface>();

        dispatch(commands, action, null!, argsJson);

        Assert.Equal(1, proxy.CallCount);
        Assert.Equal(expectedContent, proxy.LastArguments![expectedArgumentIndex]);
    }

    private static (TInterface Commands, RecordingDispatchProxy Proxy) CreateProxy<TInterface>()
        where TInterface : class
    {
        var commands = DispatchProxy.Create<TInterface, RecordingDispatchProxy>();
        return (commands, (RecordingDispatchProxy)(object)commands);
    }

    public class RecordingDispatchProxy : DispatchProxy
    {
        public int CallCount { get; private set; }

        public object?[]? LastArguments { get; private set; }

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            CallCount++;
            LastArguments = args;

            if (targetMethod == null || targetMethod.ReturnType == typeof(void))
            {
                return null;
            }

            return Activator.CreateInstance(targetMethod.ReturnType);
        }
    }
}
