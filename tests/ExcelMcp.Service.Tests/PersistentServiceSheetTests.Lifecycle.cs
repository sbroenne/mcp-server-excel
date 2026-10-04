using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceSheetTests
{
    [Fact]
    public void List_DefaultWorkbook_ReturnsDefaultSheets()
    {
        var sheets = ReadCurrentSheetNames();
        Assert.Equal(["Sheet1"], sheets);
    }

    [Fact]
    public void List_VisibleSheets_ReturnsVisibleTrue()
    {
        var result = RequireSuccess(_sheetCommands.List(_fixture.BatchToken));
        Assert.Equal("Sheet1", Assert.Single(result.Worksheets).Name);
        Assert.True(result.Worksheets[0].Visible);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets.Item["Sheet1"];
                Assert.Equal(Excel.XlSheetVisibility.xlSheetVisible, sheet.Visible);
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    [Fact]
    public void Create_UniqueName_ReturnsSuccess()
    {
        var guard = CreateSeededSheet("Guard");
        var before = CaptureSheet(guard);
        var oldNames = ReadCurrentSheetNames();
        var name = $"Create_{Guid.NewGuid():N}"[..31];
        RequireSuccess(_sheetCommands.Create(_fixture.BatchToken, name));
        _fixture.RegisterSheetForCleanup(name);
        var names = ReadCurrentSheetNames();
        Assert.Equal(oldNames.Length + 1, names.Length);
        Assert.Equal(oldNames, names.Where(item => item != name).ToArray());
        Assert.Equal(name, Assert.Single(names, item => item == name));
        Assert.All(RequireSuccess(_commands.GetValues(_fixture.BatchToken, name, "A1:C3")).Values,
            row => Assert.All(row, Assert.Null));
        Assert.Equal(before, CaptureSheet(guard));
    }

    [Fact]
    public void Rename_ExistingSheet_ReturnsSuccess()
    {
        var oldName = CreateSeededSheet("Old");
        var guard = CreateSeededSheet("Guard");
        var oldContents = CaptureSheet(oldName);
        var guardContents = CaptureSheet(guard);
        var before = ReadCurrentSheetNames();
        var newName = $"New_{Guid.NewGuid():N}"[..31];
        RequireSuccess(_sheetCommands.Rename(_fixture.BatchToken, oldName, newName));
        _fixture.RenameTrackedSheet(oldName, newName);
        Assert.Equal(before.Select(item => item == oldName ? newName : item), ReadCurrentSheetNames());
        Assert.Equal(oldContents, CaptureSheet(newName));
        Assert.Equal(guardContents, CaptureSheet(guard));
    }

    [Fact]
    public void Delete_NonActiveSheet_ReturnsSuccess()
    {
        var target = CreateSeededSheet("Delete");
        var guard = CreateSeededSheet("Guard");
        var before = ReadCurrentSheetNames();
        var retained = CaptureSheet(guard);
        RequireSuccess(_sheetCommands.Delete(_fixture.BatchToken, target));
        _fixture.ForgetSheet(target);
        Assert.Equal(before.Where(item => item != target), ReadCurrentSheetNames());
        Assert.Equal(retained, CaptureSheet(guard));
    }

    [Fact]
    public void Copy_ExistingSheet_CreatesNewSheet()
    {
        var source = CreateSeededSheet("Source");
        var guard = CreateSeededSheet("Guard");
        var before = ReadCurrentSheetNames();
        var sourceContents = CaptureSheet(source);
        var guardContents = CaptureSheet(guard);
        var target = $"Copy_{Guid.NewGuid():N}"[..31];
        RequireSuccess(_sheetCommands.Copy(_fixture.BatchToken, source, target));
        _fixture.RegisterSheetForCleanup(target);
        var after = ReadCurrentSheetNames();
        Assert.Equal(before.Length + 1, after.Length);
        Assert.Equal(before, after.Where(item => item != target).ToArray());
        Assert.Equal(target, Assert.Single(after, item => item == target));
        Assert.Equal(sourceContents, CaptureSheet(source));
        Assert.Equal(sourceContents, CaptureSheet(target));
        Assert.Equal(guardContents, CaptureSheet(guard));
    }

    [Theory]
    [InlineData("create-conflict")]
    [InlineData("rename-conflict")]
    [InlineData("rename-missing")]
    [InlineData("copy-conflict")]
    [InlineData("copy-missing")]
    [InlineData("delete-missing")]
    public void Lifecycle_RejectedTarget_PreservesCompleteSheets(string action)
    {
        var source = CreateSeededSheet("Source");
        var target = CreateSeededSheet("Guard");
        var names = ReadCurrentSheetNames();
        var sourceState = CaptureSheet(source);
        var targetState = CaptureSheet(target);
        var error = Assert.Throws<InvalidOperationException>(() =>
        {
            switch (action)
            {
                case "create-conflict": _sheetCommands.Create(_fixture.BatchToken, source); break;
                case "rename-conflict": _sheetCommands.Rename(_fixture.BatchToken, source, target); break;
                case "rename-missing": _sheetCommands.Rename(_fixture.BatchToken, "MissingSheet", "NewName"); break;
                case "copy-conflict": _sheetCommands.Copy(_fixture.BatchToken, source, target); break;
                case "copy-missing": _sheetCommands.Copy(_fixture.BatchToken, "MissingSheet", "NewName"); break;
                case "delete-missing": _sheetCommands.Delete(_fixture.BatchToken, "MissingSheet"); break;
                default: throw new ArgumentOutOfRangeException(nameof(action));
            }
        });
        var operation = action.Split('-')[0];
        Assert.Contains($"sheet.{operation} failed", error.Message, StringComparison.Ordinal);
        if (action.EndsWith("missing", StringComparison.Ordinal))
        {
            Assert.Contains("not found", error.Message, StringComparison.OrdinalIgnoreCase);
        }
        var after = ReadCurrentSheetNames();
        foreach (var unexpected in after.Except(names, StringComparer.Ordinal))
        {
            _fixture.RegisterSheetForCleanup(unexpected);
        }
        Assert.Equal(names, after);
        Assert.Equal(sourceState, CaptureSheet(source));
        Assert.Equal(targetState, CaptureSheet(target));
    }

    private string CreateSeededSheet(string prefix)
    {
        var name = $"{prefix}_{Guid.NewGuid():N}"[..Math.Min(prefix.Length + 9, 31)];
        _fixture.CreateNamedTestSheet(_fixture.BatchToken, name);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, name, "A1:B2",
            [[prefix, 7], ["second row", 11]]));
        RequireSuccess(_commands.SetFormulas(_fixture.BatchToken, name, "C2", [["=B2*2"]]));
        RequireSuccess(_commands.SetNumberFormat(_fixture.BatchToken, name, "B1:B2", "0.00"));
        var data = RequireSuccess(_commands.GetFormulas(_fixture.BatchToken, name, "A1:C3"));
        Assert.Equal(prefix, data.Values[0][0]);
        Assert.Equal("=B2*2", data.Formulas[1][2]);
        Assert.Equal(22m, Convert.ToDecimal(data.Values[1][2], System.Globalization.CultureInfo.InvariantCulture));
        return name;
    }

    public static IEnumerable<object[]> InvalidWorksheetNames()
    {
        string[] actions = ["create", "copy", "rename"];
        string[] names =
        [
            "Bad/Name", "12345678901234567890123456789012",
            "'BadName", "BadName'", "History", "", "   "
        ];
        foreach (var action in actions)
        {
            foreach (var name in names)
            {
                yield return [action, name];
            }
        }
    }

    [Theory]
    [MemberData(nameof(InvalidWorksheetNames))]
    public void Lifecycle_InvalidName_PreservesCompleteSheets(string action, string invalidName)
    {
        var source = CreateSeededSheet("Source");
        var names = ReadCurrentSheetNames();
        var state = CaptureSheet(source);
        if (!string.IsNullOrWhiteSpace(invalidName))
        {
            _fixture.ExecuteRawVerification((context, _) =>
            {
                Excel.Worksheet? sheet = null;
                try
                {
                    sheet = ComUtilities.FindSheet(context.Book, source);
                    Assert.NotNull(sheet);
                    Assert.Throws<System.Runtime.InteropServices.COMException>(() => sheet.Name = invalidName);
                    Assert.Equal(source, sheet.Name);
                }
                finally { ComUtilities.Release(ref sheet); }
            });
        }
        var error = Record.Exception(() =>
        {
            switch (action)
            {
                case "create": _sheetCommands.Create(_fixture.BatchToken, invalidName); break;
                case "copy": _sheetCommands.Copy(_fixture.BatchToken, source, invalidName); break;
                case "rename": _sheetCommands.Rename(_fixture.BatchToken, source, invalidName); break;
                default: throw new ArgumentOutOfRangeException(nameof(action));
            }
        });
        Assert.NotNull(error);
        var after = ReadCurrentSheetNames();
        foreach (var unexpected in after.Except(names, StringComparer.Ordinal))
        {
            _fixture.RegisterSheetForCleanup(unexpected);
        }
        Assert.Equal(names, after);
        Assert.Equal(state, CaptureSheet(source));
        Assert.IsType<ArgumentException>(error);
        if (string.IsNullOrWhiteSpace(invalidName))
        {
            Assert.Contains(
                $"sheet.{action} failed [InvalidInput/ArgumentException]",
                error.Message,
                StringComparison.Ordinal);
        }
        else
        {
            Assert.Contains("Worksheet names must", error.Message, StringComparison.Ordinal);
        }
    }

    public static IEnumerable<object[]> ExcelAcceptedNames()
    {
        string[] actions = ["create", "copy", "rename"];
        string[] names =
        [
            "Sales 2026", "\u88681", "\u00c9t\u00e9", "\u30c6\u30fc\u30d6\u30eb1",
            "\ud45c1", "12345", "O'Brien", " name "
        ];
        foreach (var action in actions)
        {
            foreach (var name in names)
            {
                yield return [action, name];
            }
        }
    }

    [Theory]
    [MemberData(nameof(ExcelAcceptedNames))]
    public void Lifecycle_ExcelAcceptedName_PreservesExactNameAndContents(string action, string name)
    {
        var source = CreateSeededSheet("Source");
        var guard = CreateSeededSheet("Guard");
        var before = ReadCurrentSheetNames();
        var sourceState = CaptureNativeSheet(source);
        var guardState = CaptureSheet(guard);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, source);
                Assert.NotNull(sheet);
                PersistentServiceCleanupFailures.Run(
                    () =>
                    {
                        sheet.Name = name;
                        Assert.Equal(name, sheet.Name);
                    },
                    () => sheet.Name = source);
            }
            finally { ComUtilities.Release(ref sheet); }
        });

        switch (action)
        {
            case "create": RequireSuccess(_sheetCommands.Create(_fixture.BatchToken, name)); break;
            case "copy": RequireSuccess(_sheetCommands.Copy(_fixture.BatchToken, source, name)); break;
            case "rename": RequireSuccess(_sheetCommands.Rename(_fixture.BatchToken, source, name)); break;
            default: throw new ArgumentOutOfRangeException(nameof(action));
        }
        if (action == "rename")
        {
            _fixture.RenameTrackedSheet(source, name);
            Assert.Equal(before.Select(item => item == source ? name : item), ReadCurrentSheetNames());
            Assert.Equal(sourceState, CaptureNativeSheet(name));
        }
        else
        {
            _fixture.RegisterSheetForCleanup(name);
            var after = ReadCurrentSheetNames();
            Assert.Equal(before.Length + 1, after.Length);
            Assert.Equal(before, after.Where(item => item != name).ToArray());
            Assert.Equal(name, Assert.Single(after, item => item == name));
            Assert.Equal(sourceState, CaptureNativeSheet(source));
            if (action == "copy")
            {
                Assert.Equal(sourceState, CaptureNativeSheet(name));
            }
            else
            {
                Assert.All(RequireSuccess(_commands.GetValues(_fixture.BatchToken, name, "A1:C3")).Values,
                    row => Assert.All(row, Assert.Null));
            }
        }
        Assert.Equal(guardState, CaptureSheet(guard));
    }

    private string CaptureNativeSheet(string name) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, name);
                Assert.NotNull(sheet);
                var cells = new List<object?>();
                for (var row = 1; row <= 3; row++)
                {
                    for (var column = 'A'; column <= 'C'; column++)
                    {
                        Excel.Range? cell = null;
                        try
                        {
                            cell = sheet.Range[$"{column}{row}"];
                            cells.Add(cell.Value2);
                            cells.Add(cell.Formula);
                            cells.Add(cell.NumberFormat);
                        }
                        finally { ComUtilities.Release(ref cell); }
                    }
                }
                return JsonSerializer.Serialize(cells);
            }
            finally { ComUtilities.Release(ref sheet); }
        });

    private string CaptureSheet(string name)
    {
        var result = RequireSuccess(_commands.GetFormulas(_fixture.BatchToken, name, "A1:C3"));
        var formats = RequireSuccess(_commands.GetNumberFormats(_fixture.BatchToken, name, "A1:C3"));
        Assert.Equal(3, result.Values.Count);
        Assert.All(result.Values, row => Assert.Equal(3, row.Count));
        return JsonSerializer.Serialize(new { result.Values, result.Formulas, formats.Formats });
    }

    private string[] ReadCurrentSheetNames()
    {
        var result = RequireSuccess(_sheetCommands.List(_fixture.BatchToken));
        var names = result.Worksheets.Select(sheet => sheet.Name).ToArray();
        var native = _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            try
            {
                sheets = context.Book.Worksheets;
                var nativeNames = new List<string>();
                for (var index = 1; index <= sheets.Count; index++)
                {
                    Excel.Worksheet? sheet = null;
                    try
                    {
                        sheet = (Excel.Worksheet)sheets.Item[index];
                        nativeNames.Add(sheet.Name);
                    }
                    finally { ComUtilities.Release(ref sheet); }
                }
                return nativeNames;
            }
            finally { ComUtilities.Release(ref sheets); }
        });
        Assert.Equal(native, names);
        return names;
    }
}
