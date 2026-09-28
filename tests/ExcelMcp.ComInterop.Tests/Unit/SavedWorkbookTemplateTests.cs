using System.Collections.Concurrent;
using Sbroenne.ExcelMcp.Tests.Infrastructure;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "SavedWorkbookTemplate")]
[Trait("RequiresExcel", "false")]
public sealed class SavedWorkbookTemplateTests : IDisposable
{
    private readonly string _directory =
        Path.Combine(Path.GetTempPath(), $"SavedWorkbookTemplateTests_{Guid.NewGuid():N}");

    public SavedWorkbookTemplateTests() => Directory.CreateDirectory(_directory);

    [Fact]
    public void CopyTo_CreatesOncePerBaselineAndExtension()
    {
        var creations = new ConcurrentBag<string>();
        var store = CreateStore(creations);

        store.CopyTo(Path.Combine(_directory, "first.xlsx"), "blank");
        store.CopyTo(Path.Combine(_directory, "second.xlsx"), "blank");
        store.CopyTo(Path.Combine(_directory, "macro.xlsm"), "blank");
        store.CopyTo(Path.Combine(_directory, "populated.xlsx"), "populated");

        Assert.Equal(3, creations.Count);
        Assert.Single(creations, path => Path.GetExtension(path) == ".xlsm");
        Assert.Equal(2, creations.Count(path => Path.GetExtension(path) == ".xlsx"));
    }

    [Fact]
    public void CopyTo_UsesUniqueDestinationsAndPreservesCopyIsolation()
    {
        var store = CreateStore([]);
        var first = store.CopyTo(Path.Combine(_directory, "first.xlsx"), "blank");
        var second = store.CopyTo(Path.Combine(_directory, "second.xlsx"), "blank");

        File.WriteAllText(first, "changed");

        Assert.NotEqual(first, second);
        Assert.Equal("template", File.ReadAllText(second));
    }

    [Fact]
    public async Task CopyTo_ConcurrentRequestsCreateOneTemplate()
    {
        var creationCount = 0;
        var store = new SavedWorkbookTemplateStore(
            Path.Combine(_directory, "templates"),
            path =>
            {
                Interlocked.Increment(ref creationCount);
                Thread.Sleep(50);
                File.WriteAllText(path, "template");
            });

        var destinations = Enumerable.Range(0, 16)
            .Select(index => Path.Combine(_directory, $"copy-{index}.xlsx"))
            .ToArray();

        await Task.WhenAll(destinations.Select(destination =>
            Task.Run(() => store.CopyTo(destination, "blank"))));

        Assert.Equal(1, creationCount);
        Assert.All(destinations, destination => Assert.Equal("template", File.ReadAllText(destination)));
    }

    public void Dispose()
    {
        Directory.Delete(_directory, recursive: true);
        GC.SuppressFinalize(this);
    }

    private SavedWorkbookTemplateStore CreateStore(ConcurrentBag<string> creations) =>
        new(
            Path.Combine(_directory, "templates"),
            path =>
            {
                creations.Add(path);
                File.WriteAllText(path, "template");
            });
}
