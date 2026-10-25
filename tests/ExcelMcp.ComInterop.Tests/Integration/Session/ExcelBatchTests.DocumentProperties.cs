using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Office = Microsoft.Office.Core;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

public partial class ExcelBatchTests
{
    [Fact]
    public void Execute_DocumentPropertyAdapter_ReadsWritesTypesAndPreservesOtherProperties()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, ct) =>
        {
            using var builtIn = ExcelDocumentProperties.Open(ctx.Book, true);
            using var custom = ExcelDocumentProperties.Open(ctx.Book, false);
            using var title = builtIn.Item("Title");
            using var subject = builtIn.Item("Subject");
            object? originalSubject = subject.Value;
            int builtInCount = builtIn.Count;
            int customCount = custom.Count;
            title.SetValue("Adapter title");
            Assert.Equal("Title", title.Name);
            Assert.Equal(Office.MsoDocProperties.msoPropertyTypeString, title.Type);
            Assert.False(title.LinkToContent);
            Assert.Equal("Adapter title", title.Value);

            object[] values = [17, true, new DateTime(2025, 3, 12, 10, 20, 30), "text value", 12.5d];
            Office.MsoDocProperties[] types =
            [
                Office.MsoDocProperties.msoPropertyTypeNumber,
                Office.MsoDocProperties.msoPropertyTypeBoolean,
                Office.MsoDocProperties.msoPropertyTypeDate,
                Office.MsoDocProperties.msoPropertyTypeString,
                Office.MsoDocProperties.msoPropertyTypeFloat
            ];
            for (int index = 0; index < values.Length; index++)
            {
                string name = $"AdapterType{index}";
                using var added = custom.Add(name, types[index], values[index]);
                Assert.Equal(name, added.Name);
                Assert.Equal(types[index], added.Type);
                Assert.Equal(values[index], added.Value);
                Assert.False(added.LinkToContent);
            }
            Assert.Equal(customCount + values.Length, custom.Count);
            for (int index = 0; index < values.Length; index++)
            {
                using var property = custom.Item(customCount + index + 1);
                Assert.Equal($"AdapterType{index}", property.Name);
                Assert.Equal(types[index], property.Type);
                Assert.Equal(values[index], property.Value);
            }

            using var text = custom.Item("AdapterType3");
            text.SetValue("replacement");
            Assert.Equal("replacement", text.Value);
            using var retained = custom.Item("AdapterType0");
            Assert.Equal(17, retained.Value);
            Assert.Equal(originalSubject, subject.Value);
            Assert.Equal(builtInCount, builtIn.Count);
            return 0;
        });
        AssertNoExternalInteropAssemblies();
    }

    [Fact]
    public void Execute_DocumentPropertyAdapter_ErrorsPreserveNativeCodeAndOtherProperties()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, ct) =>
        {
            using var properties = ExcelDocumentProperties.Open(ctx.Book, false);
            using var control = properties.Add("RetainedControl", Office.MsoDocProperties.msoPropertyTypeString, "retained");
            int count = properties.Count;
            var missing = Assert.ThrowsAny<COMException>(() => properties.Item("MissingProperty"));
            Assert.Equal(unchecked((int)0x80070057), missing.HResult);
            Assert.Equal(unchecked((int)0x80020009), (int)missing.Data["DispatchHResult"]!);
            Assert.NotNull(missing.InnerException);
            Assert.Equal(missing.HResult, missing.InnerException.HResult);
            Assert.Contains("'Item'", missing.Message, StringComparison.Ordinal);
            _output.WriteLine($"Missing document property: 0x{missing.HResult:X8}; dispatch 0x{(int)missing.Data["DispatchHResult"]!:X8}.");
            Assert.Equal(count, properties.Count);
            Assert.Equal("retained", control.Value);
            using var removed = properties.Add("RemovedProperty", Office.MsoDocProperties.msoPropertyTypeString, "remove");
            Assert.Equal(count + 1, properties.Count);
            removed.Delete();
            Assert.Equal(count, properties.Count);
            var deleted = Assert.ThrowsAny<COMException>(() => properties.Item("RemovedProperty"));
            Assert.Equal(missing.HResult, deleted.HResult);
            Assert.Equal("retained", control.Value);
            using var builtIn = ExcelDocumentProperties.Open(ctx.Book, true);
            using var printDate = builtIn.Item("Last Print Date");
            var unavailable = Assert.ThrowsAny<COMException>(() => printDate.Value);
            Assert.Equal(unchecked((int)0x80004005), unavailable.HResult);
            return 0;
        });
    }

    [Fact]
    public void Execute_DocumentPropertyAdapter_PreservesExactValuesAfterSaveAndReopen()
    {
        using (var batch = ExcelSession.BeginBatch(_testFileCopy!))
        {
            batch.Execute((ctx, ct) =>
            {
                using var builtIn = ExcelDocumentProperties.Open(ctx.Book, true);
                using var title = builtIn.Item("Title");
                title.SetValue("Persisted adapter title");
                using var custom = ExcelDocumentProperties.Open(ctx.Book, false);
                using var text = custom.Add("PersistedText", Office.MsoDocProperties.msoPropertyTypeString, "persisted text");
                using var number = custom.Add("PersistedNumber", Office.MsoDocProperties.msoPropertyTypeNumber, 37);
                using var flag = custom.Add("PersistedFlag", Office.MsoDocProperties.msoPropertyTypeBoolean, true);
                using var date = custom.Add("PersistedDate", Office.MsoDocProperties.msoPropertyTypeDate,
                    new DateTime(2025, 3, 12, 10, 20, 30));
                using var floating = custom.Add("PersistedFloat", Office.MsoDocProperties.msoPropertyTypeFloat, 12.5d);
                return 0;
            });
            batch.Save();
        }
        using var reopened = ExcelSession.BeginBatch(_testFileCopy!);
        reopened.Execute((ctx, ct) =>
        {
            using var builtIn = ExcelDocumentProperties.Open(ctx.Book, true);
            using var title = builtIn.Item("Title");
            Assert.Equal("Persisted adapter title", title.Value);
            using var custom = ExcelDocumentProperties.Open(ctx.Book, false);
            using var text = custom.Item("PersistedText");
            using var number = custom.Item("PersistedNumber");
            using var flag = custom.Item("PersistedFlag");
            using var date = custom.Item("PersistedDate");
            using var floating = custom.Item("PersistedFloat");
            Assert.Equal("persisted text", text.Value);
            Assert.Equal(Office.MsoDocProperties.msoPropertyTypeString, text.Type);
            Assert.Equal(37, number.Value);
            Assert.Equal(Office.MsoDocProperties.msoPropertyTypeNumber, number.Type);
            Assert.True(Assert.IsType<bool>(flag.Value));
            Assert.Equal(Office.MsoDocProperties.msoPropertyTypeBoolean, flag.Type);
            Assert.Equal(new DateTime(2025, 3, 12, 10, 20, 30), Assert.IsType<DateTime>(date.Value));
            Assert.Equal(Office.MsoDocProperties.msoPropertyTypeDate, date.Type);
            Assert.Equal(12.5d, Assert.IsType<double>(floating.Value));
            Assert.Equal(Office.MsoDocProperties.msoPropertyTypeFloat, floating.Type);
            return 0;
        });
    }

    [Fact]
    public void Execute_DocumentPropertyAdapter_RejectsOtherObjectsAndDisposedHandles()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, ct) =>
        {
            object? rawCollection = null;
            try
            {
                rawCollection = ctx.Book.CustomDocumentProperties;
                var wrongType = Assert.Throws<ArgumentException>(
                    () => new ExcelDocumentProperties.Property(rawCollection));
                Assert.Equal("property", wrongType.ParamName);
                using var collection = ExcelDocumentProperties.Open(ctx.Book, false);
                using var control = collection.Add("DisposedControl", Office.MsoDocProperties.msoPropertyTypeString, "retained");
                Assert.Equal("retained", control.Value);
                control.Dispose();
                Assert.Throws<ObjectDisposedException>(() => control.Value);
                using var retained = collection.Item("DisposedControl");
                Assert.Equal("retained", retained.Value);
                collection.Dispose();
                Assert.Throws<ObjectDisposedException>(() => collection.Count);
                Assert.Equal("retained", retained.Value);
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref rawCollection);
            }
        });
    }
}
