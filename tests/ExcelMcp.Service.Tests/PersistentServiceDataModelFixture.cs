using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Sbroenne.ExcelMcp.Tests.Infrastructure;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed class PersistentServiceDataModelFixture :
    PersistentServiceWorkbookFixture
{
    private static readonly object CreationLock = new();
    private static readonly SavedWorkbookTemplateStore Templates =
        SavedWorkbookTemplates.CreateStore(CreateTemplate);
    private static DataModelPivotTableCreationResult? _creationResult;

    public PersistentServiceDataModelFixture()
        : base(CopyTemplate, "DataModelPivotTables.xlsx")
    {
    }

    internal static DataModelPivotTableCreationResult CreationResult =>
        _creationResult
        ?? throw new InvalidOperationException(
            "The Data Model template has not been created.");

    private static void CopyTemplate(string destinationPath) =>
        Templates.CopyTo(destinationPath, "data-model-pivot-table");

    private static void CreateTemplate(string templatePath)
    {
        lock (CreationLock)
        {
            var fixture = new DataModelPivotTableFixture();
            try
            {
                fixture.InitializeAsync().GetAwaiter().GetResult();
                File.Copy(fixture.TestFilePath, templatePath, overwrite: false);
                _creationResult = fixture.CreationResult;
            }
            finally
            {
                fixture.DisposeAsync().GetAwaiter().GetResult();
            }
        }
    }
}
