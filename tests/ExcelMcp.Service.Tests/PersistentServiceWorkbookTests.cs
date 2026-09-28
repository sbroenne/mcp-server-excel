using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Workbook")]
[Trait("RequiresExcel", "true")]
public sealed partial class PersistentServiceWorkbookTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IWorkbookCommands _workbook =
        ServiceCommandProxy.Create<IWorkbookCommands>(fixture);

    [Fact]
    public void SetProtection_ProtectsAndUnprotectsWorkbook()
    {
        var batch = _fixture.BatchToken;
        var original = _workbook.GetProtection(batch);
        try
        {
            var protect = _workbook.SetProtection(batch, true);
            Assert.True(protect.Success, protect.ErrorMessage);
            var protectedState = _workbook.GetProtection(batch);
            Assert.True(protectedState.Success);
            Assert.True(protectedState.IsProtected);

            var unprotect = _workbook.SetProtection(batch, false);
            Assert.True(unprotect.Success, unprotect.ErrorMessage);
            var unprotectedState = _workbook.GetProtection(batch);
            Assert.True(unprotectedState.Success);
            Assert.False(unprotectedState.IsProtected);
        }
        finally
        {
            _workbook.SetProtection(batch, original.IsProtected);
        }
    }

    [Fact]
    public void SetViewOptions_UpdatesGridlinesAndHeadings()
    {
        var batch = _fixture.BatchToken;
        var original = _workbook.GetViewOptions(batch);
        try
        {
            var set = _workbook.SetViewOptions(
                batch,
                displayGridlines: false,
                displayHeadings: true);
            Assert.True(set.Success, set.ErrorMessage);
            var result = _workbook.GetViewOptions(batch);
            Assert.True(result.Success, result.ErrorMessage);
            Assert.False(result.DisplayGridlines);
            Assert.True(result.DisplayHeadings);
        }
        finally
        {
            _workbook.SetViewOptions(
                batch,
                original.DisplayGridlines,
                original.DisplayHeadings);
        }
    }

    [Fact]
    public void GetInfo_ReturnsActiveWorkbookMetadata()
    {
        var result = _workbook.GetInfo(_fixture.BatchToken);

        Assert.True(result.Success);
        Assert.Equal(
            Path.GetFileName(_fixture.WorkbookPath),
            result.Name);
        Assert.Equal(
            Path.GetFullPath(_fixture.WorkbookPath),
            result.FullName,
            ignoreCase: true);
        Assert.Equal("xlsx", result.Format);
        Assert.False(result.ReadOnly);
    }

    [Fact]
    public void CustomDocumentProperty_CrudRoundTrip_PreservesValue()
    {
        var batch = _fixture.BatchToken;
        var propertyName = $"AutomationTag_{Guid.NewGuid():N}";
        var deleted = false;
        try
        {
            var set = _workbook.SetDocumentProperty(
                batch,
                propertyName,
                "alpha",
                DocumentPropertyScope.Custom);
            var get = _workbook.GetDocumentProperty(
                batch,
                propertyName,
                DocumentPropertyScope.Custom);
            var list = _workbook.ListDocumentProperties(
                batch,
                includeBuiltIn: false,
                includeCustom: true);
            var delete = _workbook.DeleteDocumentProperty(
                batch,
                propertyName);
            deleted = true;

            Assert.True(set.Success);
            Assert.True(get.Success);
            Assert.Equal("alpha", get.Property.Value);
            Assert.Contains(
                list.Properties,
                property => property.Name == propertyName
                    && property.Value == "alpha"
                    && property.Scope == "custom");
            Assert.True(delete.Success);
            Assert.Throws<InvalidOperationException>(() =>
                _workbook.GetDocumentProperty(
                    batch,
                    propertyName,
                    DocumentPropertyScope.Custom));
        }
        finally
        {
            if (!deleted)
            {
                _workbook.DeleteDocumentProperty(batch, propertyName);
            }
        }
    }

    [Fact]
    public void BuiltInDocumentProperty_SetAndGet_UpdatesTitle()
    {
        var batch = _fixture.BatchToken;
        var original = _workbook.GetDocumentProperty(
            batch,
            "Title",
            DocumentPropertyScope.BuiltIn);
        try
        {
            var set = _workbook.SetDocumentProperty(
                batch,
                "Title",
                "Quarterly workbook",
                DocumentPropertyScope.BuiltIn);
            var get = _workbook.GetDocumentProperty(
                batch,
                "Title",
                DocumentPropertyScope.BuiltIn);

            Assert.True(set.Success);
            Assert.Equal("Quarterly workbook", get.Property.Value);
            Assert.Equal("built-in", get.Property.Scope);
        }
        finally
        {
            var originalTitle = original.Property.Value?.ToString();
            if (!string.IsNullOrEmpty(originalTitle))
            {
                _workbook.SetDocumentProperty(
                    batch,
                    "Title",
                    originalTitle,
                    DocumentPropertyScope.BuiltIn);
            }
        }
    }
}
