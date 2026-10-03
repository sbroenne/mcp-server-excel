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
        var original = RequireSuccess(_workbook.GetProtection(batch));
        RunWithWorkbookCleanup(() =>
        {
            RequireSuccess(_workbook.SetProtection(batch, true));
            var protectedState = RequireSuccess(_workbook.GetProtection(batch));
            Assert.True(protectedState.IsProtected);

            RequireSuccess(_workbook.SetProtection(batch, false));
            var unprotectedState = RequireSuccess(_workbook.GetProtection(batch));
            Assert.False(unprotectedState.IsProtected);
        }, () => RequireSuccess(_workbook.SetProtection(batch, original.IsProtected)));
    }

    [Fact]
    public void SetViewOptions_UpdatesGridlinesAndHeadings()
    {
        var batch = _fixture.BatchToken;
        var original = RequireSuccess(_workbook.GetViewOptions(batch));
        RunWithWorkbookCleanup(() =>
        {
            RequireSuccess(_workbook.SetViewOptions(
                batch,
                displayGridlines: false,
                displayHeadings: true));
            var result = RequireSuccess(_workbook.GetViewOptions(batch));
            Assert.False(result.DisplayGridlines);
            Assert.True(result.DisplayHeadings);
            RequireSuccess(_workbook.SetViewOptions(batch, displayGridlines: true));
            var updated = RequireSuccess(_workbook.GetViewOptions(batch));
            Assert.True(updated.DisplayGridlines);
            Assert.True(updated.DisplayHeadings);
        }, () =>
            RequireSuccess(_workbook.SetViewOptions(
                batch,
                original.DisplayGridlines,
                original.DisplayHeadings)));
    }

    [Fact]
    public void GetInfo_ReturnsActiveWorkbookMetadata()
    {
        var result = RequireSuccess(_workbook.GetInfo(_fixture.BatchToken));

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
        var created = false;
        var deleted = false;
        RunWithWorkbookCleanup(() =>
        {
            RequireSuccess(_workbook.SetDocumentProperty(
                batch,
                propertyName,
                "alpha",
                DocumentPropertyScope.Custom));
            created = true;
            var get = RequireSuccess(_workbook.GetDocumentProperty(
                batch,
                propertyName,
                DocumentPropertyScope.Custom));
            Assert.Equal("alpha", get.Property.Value);
            Assert.Equal(propertyName, get.Property.Name);
            Assert.Equal("custom", get.Property.Scope);
            var list = RequireSuccess(_workbook.ListDocumentProperties(
                batch,
                includeBuiltIn: false,
                includeCustom: true));

            Assert.Single(
                list.Properties,
                property => property.Name == propertyName
                    && property.Value == "alpha"
                    && property.Scope == "custom");
            RequireSuccess(_workbook.SetDocumentProperty(
                batch, propertyName, "beta", DocumentPropertyScope.Custom));
            var updated = RequireSuccess(_workbook.GetDocumentProperty(
                batch, propertyName, DocumentPropertyScope.Custom));
            Assert.Equal("beta", updated.Property.Value);
            RequireSuccess(_workbook.DeleteDocumentProperty(batch, propertyName));
            deleted = true;
            Assert.DoesNotContain(
                RequireSuccess(_workbook.ListDocumentProperties(
                    batch, includeBuiltIn: false, includeCustom: true)).Properties,
                property => property.Name == propertyName);
            Assert.Throws<InvalidOperationException>(() =>
                _workbook.GetDocumentProperty(
                    batch,
                    propertyName,
                    DocumentPropertyScope.Custom));
        }, () =>
        {
            if (created && !deleted)
            {
                RequireSuccess(_workbook.DeleteDocumentProperty(batch, propertyName));
            }
        });
    }

    [Fact]
    public void BuiltInDocumentProperty_SetAndGet_UpdatesTitle()
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_workbook.SetDocumentProperty(
            batch, "Title", "Quarterly workbook", DocumentPropertyScope.BuiltIn));
        var get = RequireSuccess(_workbook.GetDocumentProperty(
            batch, "Title", DocumentPropertyScope.BuiltIn));
        Assert.Equal("Quarterly workbook", get.Property.Value);
        Assert.Equal("built-in", get.Property.Scope);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MissingDocumentProperty_PreservesExistingProperty(bool delete)
    {
        var batch = _fixture.BatchToken;
        var name = $"Retained_{Guid.NewGuid():N}";
        var missing = $"Missing_{Guid.NewGuid():N}";
        RequireSuccess(_workbook.SetDocumentProperty(batch, name, "retained", DocumentPropertyScope.Custom));
        RunWithWorkbookCleanup(() =>
        {
            var before = System.Text.Json.JsonSerializer.Serialize(
                RequireSuccess(_workbook.ListDocumentProperties(
                    batch, includeBuiltIn: false, includeCustom: true)).Properties);
            var error = Assert.Throws<InvalidOperationException>(() =>
            {
                if (delete)
                {
                    _workbook.DeleteDocumentProperty(batch, missing);
                }
                else
                {
                    _workbook.GetDocumentProperty(batch, missing, DocumentPropertyScope.Custom);
                }
            });
            Assert.Contains(missing, error.Message, StringComparison.Ordinal);
            Assert.Contains("not found", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(before, System.Text.Json.JsonSerializer.Serialize(
                RequireSuccess(_workbook.ListDocumentProperties(
                    batch, includeBuiltIn: false, includeCustom: true)).Properties));
            Assert.Equal("retained", RequireSuccess(
                _workbook.GetDocumentProperty(batch, name, DocumentPropertyScope.Custom)).Property.Value);
        }, () => RequireSuccess(_workbook.DeleteDocumentProperty(batch, name)));
    }

    private static void RunWithWorkbookCleanup(Action test, Action cleanup)
    {
        Exception? failure = null;
        try
        {
            test();
        }
        catch (Exception ex)
        {
            failure = ex;
        }
        finally
        {
            try
            {
                cleanup();
            }
            catch (Exception cleanupFailure)
            {
                failure = PersistentServiceCleanupFailures.Combine(failure, cleanupFailure);
            }
        }
        if (failure is not null)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw(failure);
        }
    }
}
