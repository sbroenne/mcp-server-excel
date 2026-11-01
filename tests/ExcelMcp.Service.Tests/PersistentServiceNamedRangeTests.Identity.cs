using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceNamedRangeTests
{
    [Theory]
    [InlineData("leading-space", false)]
    [InlineData("leading-space", true)]
    [InlineData("trailing-space", false)]
    [InlineData("trailing-space", true)]
    [InlineData("quoted-scope", false)]
    [InlineData("quoted-scope", true)]
    [InlineData("unqualified-local", false)]
    [InlineData("unqualified-local", true)]
    [InlineData("different-case", false)]
    [InlineData("different-case", true)]
    [InlineData("sharp-s", false)]
    [InlineData("sharp-s", true)]
    [InlineData("canonical-accent", false)]
    [InlineData("canonical-accent", true)]
    public async Task Create_ExistingNameAlias_PreservesDefinitionMetadataAndSavedState(
        string alias,
        bool rejectedFormula)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateNamedTestSheet(batch, $"Identity_{Guid.NewGuid():N}"[..25]);
        var prefix = CreateUniqueNamedRangeName();
        var leafName = alias switch
        {
            "sharp-s" => prefix + "Straße",
            "canonical-accent" => prefix + "Caf\u00e9",
            _ => prefix
        };
        var scoped = alias is "quoted-scope" or "unqualified-local";
        var existingName = AddIdentityName(
            scoped ? $"{sheetName}!{leafName}" : leafName,
            visible: false);
        var requestedName = alias switch
        {
            "leading-space" => $" {leafName}",
            "trailing-space" => $"{leafName} ",
            "quoted-scope" => $"'{sheetName}'!{leafName}",
            "unqualified-local" => leafName,
            "different-case" => leafName.ToLowerInvariant(),
            "sharp-s" => prefix + "STRASSE",
            "canonical-accent" => prefix + "Cafe\u0301",
            _ => throw new ArgumentOutOfRangeException(nameof(alias))
        };
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [[17]]).Success);
        ActivateIdentitySheet(sheetName);
        _fixture.ExecuteRawVerification((ctx, _) => ctx.Book.Save());
        var before = CaptureIdentityNames();
        Assert.True(before.Saved);
        var original = Assert.Single(before.Names, item => item.Name == existingName);
        Assert.False(original.Visible);
        Assert.Equal("Retain this name's metadata.", original.Comment);
        var reference = rejectedFormula
            ? $"=SUM(MissingIdentityTable_{Guid.NewGuid():N}[Value])"
            : "=1.25";

        var failed = await _fixture.SendForFailureAsync("namedrange.create", new
        {
            name = requestedName,
            reference
        });

        Assert.False(failed.Success);
        var after = CaptureIdentityNames();
        Assert.Equal(before.Names, after.Names);
        Assert.True(after.Saved, "Rejecting an existing name must leave the saved workbook unchanged.");
        var values = _commands.GetValues(batch, sheetName, "A1");
        Assert.True(values.Success, values.ErrorMessage);
        Assert.Equal(17d, Convert.ToDouble(Assert.Single(Assert.Single(values.Values)),
            CultureInfo.InvariantCulture));

        await _fixture.SaveAndReopenAsync();
        Assert.Equal(before.Names, CaptureIdentityNames().Names);
    }

    [Theory]
    [InlineData(false, "Scope B")]
    [InlineData(true, "Scope B")]
    [InlineData(true, "O'Brien")]
    [InlineData(true, "Scope!Bang")]
    public async Task Create_SameLeafInDifferentScope_PreservesExistingName(
        bool existingIsLocal,
        string secondSheetPrefix)
    {
        var batch = _fixture.BatchToken;
        var firstSheet = _fixture.CreateNamedTestSheet(batch, $"ScopeA_{Guid.NewGuid():N}"[..25]);
        var secondSheet = _fixture.CreateNamedTestSheet(batch, $"{secondSheetPrefix}_{Guid.NewGuid():N}"[..25]);
        var leafName = CreateUniqueNamedRangeName();
        var existingName = AddIdentityName(
            existingIsLocal ? $"{firstSheet}!{leafName}" : leafName,
            visible: false);
        var requestedName = existingIsLocal
            ? $"'{secondSheet.Replace("'", "''", StringComparison.Ordinal)}'!{leafName}"
            : $"{firstSheet}!{leafName}";
        ActivateIdentitySheet(firstSheet);
        var before = CaptureIdentityNames();

        Assert.True(_parameterCommands.Create(batch, requestedName, "=1.25").Success);
        var after = CaptureIdentityNames();
        var added = Assert.Single(after.Names, item =>
            !before.Names.Any(previous => previous.Name == item.Name));
        _fixture.RegisterNamedRangeForCleanup(added.Name);

        Assert.Equal(requestedName, added.Name);
        Assert.Equal("=1.25", added.Reference);
        Assert.Equal("=1.25", added.R1C1Reference);
        Assert.True(added.Visible);
        Assert.Equal(before.Names, after.Names.Where(item => item.Name != added.Name));
        Assert.Equal(Assert.Single(before.Names, item => item.Name == existingName),
            Assert.Single(after.Names, item => item.Name == existingName));

        await _fixture.SaveAndReopenAsync();
        Assert.Equal(after.Names, CaptureIdentityNames().Names);
    }

    [Theory]
    [InlineData("", "\u2003")]
    [InlineData("I", "\u0131")]
    [InlineData("\u0130", "i")]
    [InlineData("\uff41", "a")]
    [InlineData("\u017f", "s")]
    public async Task Create_DistinctUnicodeName_PreservesExistingDefinition(
        string existingSuffix,
        string requestedSuffix)
    {
        var prefix = CreateUniqueNamedRangeName();
        AddIdentityName(prefix + existingSuffix, visible: false);
        var requestedName = prefix + requestedSuffix;
        var before = CaptureIdentityNames();

        Assert.True(_parameterCommands.Create(_fixture.BatchToken, requestedName, "=1.25").Success);
        var after = CaptureIdentityNames();
        var added = Assert.Single(after.Names, item => item.Name == requestedName);
        _fixture.RegisterNamedRangeForCleanup(added.Name);

        Assert.Equal("=1.25", added.Reference);
        Assert.Equal("=1.25", added.R1C1Reference);
        Assert.True(added.Visible);
        Assert.Equal(before.Names, after.Names.Where(item => item.Name != requestedName));

        await _fixture.SaveAndReopenAsync();
        Assert.Equal(after.Names, CaptureIdentityNames().Names);
    }

    private string AddIdentityName(string requestedName, bool visible)
    {
        var actualName = _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Names? names = null;
            Excel.Name? name = null;
            try
            {
                names = ctx.Book.Names;
                name = names.Add(Name: requestedName, RefersTo: "=7");
                name.Visible = visible;
                name.Comment = "Retain this name's metadata.";
                return name.Name;
            }
            finally
            {
                ComUtilities.Release(ref name);
                ComUtilities.Release(ref names);
            }
        });
        _fixture.RegisterNamedRangeForCleanup(actualName);
        return actualName;
    }

    private void ActivateIdentitySheet(string sheetName) =>
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets.Item[sheetName];
                sheet.Activate();
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });

    private (bool Saved, IdentityNameState[] Names) CaptureIdentityNames() =>
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Names? names = null;
            var result = new List<IdentityNameState>();
            try
            {
                names = ctx.Book.Names;
                for (var index = 1; index <= names.Count; index++)
                {
                    Excel.Name? name = null;
                    try
                    {
                        name = names.Item(index);
                        var reference = Assert.IsType<string>((object?)name.RefersTo);
                        var r1c1Reference = Assert.IsType<string>((object?)name.RefersToR1C1);
                        result.Add(new IdentityNameState(name.Name, reference,
                            r1c1Reference, name.Visible, name.Comment));
                    }
                    finally
                    {
                        ComUtilities.Release(ref name);
                    }
                }
                return (ctx.Book.Saved, result.OrderBy(item => item.Name, StringComparer.Ordinal).ToArray());
            }
            finally
            {
                ComUtilities.Release(ref names);
            }
        });

    private sealed record IdentityNameState(
        string Name,
        string Reference,
        string R1C1Reference,
        bool Visible,
        string Comment);
}
