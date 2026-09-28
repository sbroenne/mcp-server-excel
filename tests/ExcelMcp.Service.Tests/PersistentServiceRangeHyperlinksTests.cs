using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
public sealed partial class PersistentServiceRangeHyperlinksTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>;