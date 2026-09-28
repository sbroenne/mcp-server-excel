using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Tests.Infrastructure;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

[Collection("Sequential")]
[Trait("Category", "Integration")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "SavedWorkbookTemplate")]
[Trait("RequiresExcel", "true")]
public sealed class SavedWorkbookTemplateExcelTests
{
    [Fact]
    public void CopyOpenMutateSaveReopen_IsolatedAndUsesOwnedExcelPath()
    {
        var directory = Path.Combine(
            Path.GetTempPath(),
            $"SavedWorkbookTemplateExcelTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            var template = Path.Combine(directory, "template.xlsx");
            using (var manager = new SessionManager())
            {
                var sessionId = manager.CreateSessionForNewFile(template, show: false);
                manager.CloseSession(sessionId, save: true);
            }

            var store = new SavedWorkbookTemplateStore(
                Path.Combine(directory, "templates"),
                path => File.Copy(template, path));
            var tracked = new List<ExcelProcessIdentity>();
            void OnTracked(ExcelProcessIdentity identity) => tracked.Add(identity);

            SessionManager.ExcelProcessIdentityTracked += OnTracked;
            try
            {
                var first = store.CopyTo(Path.Combine(directory, "first.xlsx"), "blank");
                var second = store.CopyTo(Path.Combine(directory, "second.xlsx"), "blank");
                Assert.Empty(tracked);

                using (var batch = ExcelSession.BeginBatch(first))
                {
                    batch.Execute((context, _) => context.Book.Worksheets[1].Range["A1"].Value2 = "changed");
                    batch.Save();
                }

                Assert.NotEmpty(tracked);

                using var firstReopened = ExcelSession.BeginBatch(first);
                using var secondReopened = ExcelSession.BeginBatch(second);
                var firstValue = firstReopened.Execute(
                    (context, _) => (string?)context.Book.Worksheets[1].Range["A1"].Value2);
                var secondValue = secondReopened.Execute(
                    (context, _) => (string?)context.Book.Worksheets[1].Range["A1"].Value2);

                Assert.Equal("changed", firstValue);
                Assert.Null(secondValue);
            }
            finally
            {
                SessionManager.ExcelProcessIdentityTracked -= OnTracked;
            }
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }
}
