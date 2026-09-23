using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Vba;

[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "VBA")]
[Trait("Speed", "Medium")]
public sealed class FailurePrerequisiteTests : IClassFixture<VbaTestsFixture>
{
    private readonly VbaTestsFixture _fixture;

    public FailurePrerequisiteTests(VbaTestsFixture fixture) => _fixture = fixture;

    [Theory]
    [InlineData("view")]
    [InlineData("update")]
    [InlineData("delete")]
    public void MissingModule_HasNotFoundCategory(string action)
    {
        using var batch = ExcelSession.BeginBatch(_fixture.CreateTestFile());
        var commands = new VbaCommands();
        var error = Assert.ThrowsAny<InvalidOperationException>(() =>
        {
            switch (action)
            {
                case "view": commands.View(batch, "MissingModule"); break;
                case "update": commands.Update(batch, "MissingModule", "Sub Probe()\nEnd Sub"); break;
                case "delete": commands.Delete(batch, "MissingModule"); break;
            }
        });

        Assert.Contains("MissingModule", error.Message, StringComparison.Ordinal);
        Assert.Equal(OperationFailureCategory.NotFound, Assert.IsType<OperationFailureException>(error).ErrorCategory);
        Assert.DoesNotContain(commands.List(batch).Scripts, module => module.Name == "MissingModule");
    }

    [Fact]
    public void Import_ExistingModule_HasConflictCategory()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.CreateTestFile());
        var commands = new VbaCommands();
        const string code = "Sub Probe()\nEnd Sub";
        commands.Import(batch, "ExistingModule", code);
        var error = Assert.ThrowsAny<InvalidOperationException>(() =>
            commands.Import(batch, "ExistingModule", "Sub Replacement()\nEnd Sub"));

        Assert.Equal(OperationFailureCategory.Conflict, Assert.IsType<OperationFailureException>(error).ErrorCategory);
        var savedCode = commands.View(batch, "ExistingModule").Code;
        Assert.Contains("Sub Probe()", savedCode, StringComparison.Ordinal);
        Assert.DoesNotContain("Replacement", savedCode, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    public void Run_EmptyProcedure_IsInvalidInput(string procedure)
    {
        using var batch = ExcelSession.BeginBatch(_fixture.CreateTestFile());
        var error = Assert.Throws<ArgumentException>(() => new VbaCommands().Run(batch, procedure, null));
        Assert.Equal("procedureName", error.ParamName);
    }

    [Fact]
    public void UnsupportedWorkbookFormat_IsInvalidInputWhileListRemainsEmpty()
    {
        var path = Path.Join(_fixture.TempDir, $"{Guid.NewGuid():N}.xlsx");
        using var manager = new SessionManager();
        var session = manager.CreateSessionForNewFile(path, show: false);
        manager.CloseSession(session, save: true);
        using var batch = ExcelSession.BeginBatch(path);
        var commands = new VbaCommands();

        var list = commands.List(batch);
        Assert.True(list.Success);
        Assert.Null(list.ErrorMessage);
        Assert.Empty(list.Scripts);

        Action[] operations =
        [
            () => commands.Import(batch, "Probe", "Sub Probe()\nEnd Sub"),
            () => commands.View(batch, "Probe"),
            () => commands.Update(batch, "Probe", "Sub Probe()\nEnd Sub"),
            () => commands.Delete(batch, "Probe")
        ];
        foreach (var operation in operations)
        {
            var error = Assert.Throws<OperationFailureException>(operation);
            Assert.Equal(OperationFailureCategory.InvalidInput, error.ErrorCategory);
            Assert.Contains(".xlsm", error.Message, StringComparison.Ordinal);
        }

        Assert.Throws<ArgumentException>(() => commands.Run(batch, "Probe", null));
    }

    [Fact]
    [Trait("Feature", "DataModel")]
    public void Evaluate_MissingModel_HasPrerequisiteCategory()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.CreateTestFile());
        var error = Assert.ThrowsAny<InvalidOperationException>(() =>
            new DataModelCommands().Evaluate(batch, "EVALUATE ROW(\"Value\", 1)"));

        Assert.Contains("Data Model", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(OperationFailureCategory.Prerequisite, Assert.IsType<OperationFailureException>(error).ErrorCategory);
    }
}
