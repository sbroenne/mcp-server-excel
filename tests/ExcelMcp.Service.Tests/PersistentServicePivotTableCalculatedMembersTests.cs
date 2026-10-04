using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for PivotTable calculated members operations.
/// Calculated members are OLAP-only features that work with Data Model PivotTables.
/// </summary>
[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PivotTables")]
[Trait("Speed", "Slow")]
public class PersistentServicePivotTableCalculatedMembersTests(
    PersistentServiceDataModelFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceDataModelFixture>
{
    private readonly IPersistentPivotTableCommands _pivotCommands =
        fixture.CreateCommands<IPersistentPivotTableCommands>();
    private readonly DataModelPivotTableCreationResult _creationResult =
        PersistentServiceDataModelFixture.CreationResult;

    /// <summary>
    /// Tests listing calculated members on an OLAP PivotTable without any calculated members.
    /// </summary>
    [Fact]
    public void ListCalculatedMembers_OlapPivotTableNoMembers_ReturnsEmptyList()
    {
        // Arrange
        Assert.True(_creationResult.Success, "Data Model fixture must be created successfully");

        // Act
        var batch = _fixture.BatchToken;
        var result = _pivotCommands.ListCalculatedMembers(batch, "DataModelPivot");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.CalculatedMembers);
        Assert.Empty(result.CalculatedMembers);
    }

    /// <summary>
    /// Tests that an invalid measure formula returns a public error response.
    /// </summary>
    [Fact]
    public void CreateCalculatedMember_InvalidMeasureFormula_ReturnsError()
    {
        Assert.True(_creationResult.Success, "Data Model fixture must be created successfully");

        var batch = _fixture.BatchToken;
        CreateGuardSet();
        var before = ReadState();
        var result = _pivotCommands.CreateCalculatedMember(
            batch,
            "DataModelPivot",
            "[Measures].[InvalidMeasure]",
            "NOT VALID MDX(",
            CalculatedMemberType.Measure);

        Assert.False(result.Success);
        Assert.Contains(
            "Invalid formula syntax for calculated Measure",
            result.ErrorMessage);
        Assert.Equal(before, ReadState());
    }

    /// <summary>
    /// Tests creating and then deleting a calculated member.
    /// </summary>
    [Fact]
    public void DeleteCalculatedMember_ExistingSet_RemovesSet()
    {
        // Arrange
        Assert.True(_creationResult.Success, "Data Model fixture must be created successfully");

        var batch = _fixture.BatchToken;

        const string setName = "[ToBeDeletedSet]";
        CreateGuardSet();
        var before = ReadState();

        var createResult = _pivotCommands.CreateCalculatedMember(
            batch,
            "DataModelPivot",
            setName,
            "'{[RegionalSalesTable].[Region].Members}'",
            CalculatedMemberType.Set);

        Assert.True(
            createResult.Success,
            $"Calculated set creation failed: {createResult.ErrorMessage}");

        _fixture.RegisterCalculatedMemberForCleanup(
            "DataModelPivot",
            setName);

        // Verify it exists
        var listBefore = _pivotCommands.ListCalculatedMembers(batch, "DataModelPivot");
        Assert.True(listBefore.Success, listBefore.ErrorMessage);
        Assert.Contains(
            listBefore.CalculatedMembers,
            member => member.Name == setName);

        var deleteResult = _pivotCommands.DeleteCalculatedMember(
            batch,
            "DataModelPivot",
            setName);

        Assert.True(deleteResult.Success, $"Delete failed: {deleteResult.ErrorMessage}");
        _fixture.ForgetCalculatedMember(
            "DataModelPivot",
            setName);

        // Verify it's gone
        var listAfter = _pivotCommands.ListCalculatedMembers(batch, "DataModelPivot");
        Assert.True(listAfter.Success, listAfter.ErrorMessage);
        Assert.DoesNotContain(
            listAfter.CalculatedMembers,
            member => member.Name == setName);
        Assert.Equal(before, ReadState());
    }

    /// <summary>
    /// Tests deleting a non-existent calculated member returns appropriate error.
    /// </summary>
    [Fact]
    public void DeleteCalculatedMember_NonExistentMember_ReturnsError()
    {
        // Arrange
        Assert.True(_creationResult.Success, "Data Model fixture must be created successfully");

        // Act
        var batch = _fixture.BatchToken;
        CreateGuardSet();
        var before = ReadState();
        var result = _pivotCommands.DeleteCalculatedMember(batch, "DataModelPivot", "[Measures].[NonExistent]");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage);
        Assert.Contains("list-calculated-members", result.ErrorMessage);
        Assert.Equal(before, ReadState());
    }

    /// <summary>
    /// Tests creating a supported named set on the local Data Model cube.
    /// </summary>
    [Fact]
    public void CreateCalculatedMember_ValidSet_ReturnsSuccess()
    {
        Assert.True(_creationResult.Success, "Data Model fixture must be created successfully");

        var batch = _fixture.BatchToken;
        const string setName = "[TopRegionsSet]";
        var result = _pivotCommands.CreateCalculatedMember(
            batch,
            "DataModelPivot",
            setName,
            "'{[RegionalSalesTable].[Region].Members}'",
            CalculatedMemberType.Set);

        Assert.True(
            result.Success,
            $"Calculated set creation failed: {result.ErrorMessage}");
        _fixture.RegisterCalculatedMemberForCleanup("DataModelPivot", setName);
        Assert.Equal(setName, result.Name);
        Assert.Equal(CalculatedMemberType.Set, result.Type);
        Assert.True(result.IsValid);

        var listResult = _pivotCommands.ListCalculatedMembers(
            batch,
            "DataModelPivot");
        Assert.True(listResult.Success, listResult.ErrorMessage);
        Assert.Contains(
            listResult.CalculatedMembers,
            member =>
                member.Name == setName &&
                member.Type == CalculatedMemberType.Set &&
                member.IsValid);
    }

    private void CreateGuardSet()
    {
        const string name = "[RetainedRegionsSet]";
        var result = RequireSuccess(_pivotCommands.CreateCalculatedMember(_fixture.BatchToken,
            "DataModelPivot", name, "'{[RegionalSalesTable].[Region].Members}'", CalculatedMemberType.Set));
        _fixture.RegisterCalculatedMemberForCleanup("DataModelPivot", name);
        Assert.Equal(name, result.Name);
        var member = Assert.Single(RequireSuccess(_pivotCommands.ListCalculatedMembers(
            _fixture.BatchToken, "DataModelPivot")).CalculatedMembers, member => member.Name == name);
        Assert.Equal(CalculatedMemberType.Set, member.Type);
        Assert.True(member.IsValid);
        Assert.False(string.IsNullOrWhiteSpace(member.Formula));
    }

    private string ReadState()
    {
        var batch = _fixture.BatchToken;
        var members = RequireSuccess(_pivotCommands.ListCalculatedMembers(batch, "DataModelPivot"));
        var fields = RequireSuccess(_pivotCommands.ListFields(batch, "DataModelPivot"));
        var data = RequireSuccess(_pivotCommands.GetData(batch, "DataModelPivot"));
        Assert.NotEmpty(data.Values);
        return JsonSerializer.Serialize(new { members.CalculatedMembers, fields.Fields, fields.ValueFields, data });
    }
}
