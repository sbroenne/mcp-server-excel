using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceNamedRangeTests
{
    [Fact]
    public void List_EmptyWorkbook_ReturnsEmptyList()
    {
        // Arrange - Use the shared fixture file
        var batch = _fixture.BatchToken;

        // Act
        var namedRanges = _parameterCommands.List(batch);

        // Assert - List should return without error
        Assert.NotNull(namedRanges);
        Assert.True(namedRanges.Success);
        Assert.Empty(namedRanges.NamedRanges);
    }
    /// <inheritdoc/>

    [Fact]
    public void Create_ValidNameAndReference_ReturnsSuccess()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateNamedTestSheet(batch, $"Named range {Guid.NewGuid():N}"[..27]);
        var paramName = CreateUniqueNamedRangeName();
        var cellRef = $"'{sheetName}'!$A$1";
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [["Named value"]]).Success);

        // Act
        Assert.True(_parameterCommands.Create(batch, paramName, cellRef).Success);
        _fixture.RegisterNamedRangeForCleanup(paramName);

        // Assert - Verify the parameter was actually created by listing parameters
        var namedRanges = _parameterCommands.List(batch);
        Assert.True(namedRanges.Success);
        var created = Assert.Single(namedRanges.NamedRanges, p => p.Name == paramName);
        Assert.Equal($"={cellRef}", created.RefersTo);
        Assert.Equal("Named value", created.Value);
        Assert.Equal("String", created.ValueType);
        Assert.Equal(1, created.CellCount);
    }
    /// <inheritdoc/>

    [Fact]
    public void Delete_ExistingParameter_ReturnsSuccess()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var paramName = CreateUniqueNamedRangeName();
        var cellRef = $"'{sheetName}'!A1";

        // Create parameter first
        Assert.True(_parameterCommands.Create(batch, paramName, cellRef).Success);
        _fixture.RegisterNamedRangeForCleanup(paramName);
        var untouchedName = CreateUniqueNamedRangeName();
        Assert.True(_parameterCommands.Create(batch, untouchedName, $"'{sheetName}'!$B$1").Success);
        _fixture.RegisterNamedRangeForCleanup(untouchedName);
        Assert.True(_commands.SetValues(batch, sheetName, "A1:B1", [["Retained", "Untouched"]]).Success);
        var before = _parameterCommands.List(batch);
        Assert.True(before.Success, before.ErrorMessage);
        Assert.Single(before.NamedRanges, p => p.Name == paramName);
        var untouchedBefore = Assert.Single(before.NamedRanges, p => p.Name == untouchedName);

        // Delete the parameter
        Assert.True(_parameterCommands.Delete(batch, paramName).Success);
        _fixture.ForgetNamedRange(paramName);

        // Assert - Verify the parameter was actually deleted by checking it's not in the list
        var namedRanges = _parameterCommands.List(batch);
        Assert.True(namedRanges.Success);
        Assert.DoesNotContain(namedRanges.NamedRanges, p => p.Name == paramName);
        var untouched = Assert.Single(namedRanges.NamedRanges);
        Assert.Equal(untouchedName, untouched.Name);
        Assert.Equal(untouchedBefore.RefersTo, untouched.RefersTo);
        Assert.Equal("Untouched", untouched.Value);
        var cells = _commands.GetValues(batch, sheetName, "A1:B1");
        Assert.True(cells.Success, cells.ErrorMessage);
        Assert.Equal(["Retained", "Untouched"], Assert.Single(cells.Values));
    }

    [Fact]
    public void Update_ExistingReference_MovesNamedTargetWithoutChangingCells()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateNamedTestSheet(batch, $"Named update {Guid.NewGuid():N}"[..27]);
        var name = CreateUniqueNamedRangeName();
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [["Original"]]).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "D3", [["New target"]]).Success);
        Assert.True(_parameterCommands.Create(batch, name, $"'{sheetName}'!$A$1").Success);
        _fixture.RegisterNamedRangeForCleanup(name);
        Assert.Equal("Original", _parameterCommands.Read(batch, name).Value);

        var updated = _parameterCommands.Update(batch, name, $"'{sheetName}'!$D$3");

        Assert.True(updated.Success, updated.ErrorMessage);
        var read = _parameterCommands.Read(batch, name);
        Assert.Equal(name, read.Name);
        Assert.Equal($"='{sheetName}'!$D$3", read.RefersTo);
        Assert.Equal("New target", read.Value);
        Assert.Equal("String", read.ValueType);
        var original = _commands.GetValues(batch, sheetName, "A1");
        Assert.True(original.Success, original.ErrorMessage);
        Assert.Equal("Original", Assert.Single(Assert.Single(original.Values)));
        var destination = _commands.GetValues(batch, sheetName, "D3");
        Assert.True(destination.Success, destination.ErrorMessage);
        Assert.Equal("New target", Assert.Single(Assert.Single(destination.Values)));
    }
}
