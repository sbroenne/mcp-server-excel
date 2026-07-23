// <copyright file="ScreenshotTestsFixture.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Shared fixture for ScreenshotCommandsTests.
/// Each test creates its own file with data via CreateTestFile().
/// </summary>
public class ScreenshotTestsFixture : IAsyncLifetime
{
    private readonly string _tempDir;

    /// <summary>
    /// Initializes a new instance of the <see cref="ScreenshotTestsFixture"/> class.
    /// </summary>
    public ScreenshotTestsFixture()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"ScreenshotTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <summary>
    /// Creates a unique test file for a test.
    /// </summary>
    public string CreateTestFile([CallerMemberName] string testName = "")
    {
        var fileName = $"{testName}_{Guid.NewGuid():N}.xlsx";
        var filePath = Path.Combine(_tempDir, fileName);

        using var manager = new SessionManager();
        var sessionId = manager.CreateSessionForNewFile(filePath, show: false);
        manager.CloseSession(sessionId, save: true);

        return filePath;
    }

    /// <inheritdoc/>
    public Task InitializeAsync() => Task.CompletedTask;

    /// <inheritdoc/>
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
            // Ignore cleanup errors
        }

        return Task.CompletedTask;
    }
}

