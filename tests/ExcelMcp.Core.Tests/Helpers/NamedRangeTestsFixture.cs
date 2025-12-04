// <copyright file="NamedRangeTestsFixture.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Fixture that creates ONE NamedRange test file per test CLASS.
/// Each test uses unique named range names for isolation.
/// - Created ONCE before any tests run (~3-5s)
/// - Shared by all tests in the class
/// - Each test creates its own named ranges with unique names
/// - Reduces file creation overhead by ~95%
/// </summary>
public class NamedRangeTestsFixture : IAsyncLifetime
{
    private readonly string _tempDir;
    private readonly FileCommands _fileCommands = new();
    private int _cellCounter;

    /// <summary>
    /// Path to the shared test file
    /// </summary>
    public string TestFilePath { get; private set; } = null!;

    /// <summary>
    /// Results of fixture creation (exposed for validation)
    /// </summary>
    public FixtureCreationResult CreationResult { get; private set; } = null!;

    /// <summary>
    /// Initializes a new instance of the fixture
    /// </summary>
    public NamedRangeTestsFixture()
    {
        _tempDir = Path.Join(Path.GetTempPath(), $"NamedRangeTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <summary>
    /// Gets a unique named range name for test isolation.
    /// Named range names are limited to 255 chars (Excel limit).
    /// </summary>
    public string GetUniqueNamedRangeName([System.Runtime.CompilerServices.CallerMemberName] string testName = "")
    {
        // Create a unique name that fits within Excel's 255 character limit
        var uniqueId = Guid.NewGuid().ToString("N")[..8];
        var prefix = $"NR_{uniqueId}_";
        var maxNameLength = 255 - prefix.Length;
        var shortName = testName.Length > maxNameLength ? testName[..maxNameLength] : testName;
        return $"{prefix}{shortName}";
    }

    /// <summary>
    /// Gets a unique cell reference for the named range (to avoid conflicts).
    /// Uses incrementing row numbers to avoid conflicts.
    /// </summary>
    public string GetUniqueCellReference()
    {
        // Use incrementing counter for unique cell references
        var counter = Interlocked.Increment(ref _cellCounter);
        var col = ((counter - 1) % 26) + 1; // A-Z
        var row = ((counter - 1) / 26) + 1; // 1, 2, 3...
        var colLetter = (char)('A' + col - 1);
        return $"Sheet1!{colLetter}{row}";
    }

    /// <summary>
    /// Called ONCE before any tests in the class run.
    /// Creates a workbook that tests will add named ranges to.
    /// </summary>
    public Task InitializeAsync()
    {
        var sw = Stopwatch.StartNew();

        TestFilePath = Path.Join(_tempDir, "NamedRangeTest.xlsx");
        CreationResult = new FixtureCreationResult();

        try
        {
            _fileCommands.CreateEmpty(TestFilePath);
            CreationResult.FileCreated = true;

            sw.Stop();
            CreationResult.Success = true;
            CreationResult.CreationTimeSeconds = sw.Elapsed.TotalSeconds;
        }
        catch (Exception ex)
        {
            CreationResult.Success = false;
            CreationResult.ErrorMessage = ex.Message;
            sw.Stop();
            throw;
        }

        return Task.CompletedTask;
    }

    /// <summary>
    /// Called ONCE after all tests in the class complete.
    /// </summary>
    public Task DisposeAsync()
    {
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }
        catch
        {
            // Cleanup is best-effort
        }

        return Task.CompletedTask;
    }
}
