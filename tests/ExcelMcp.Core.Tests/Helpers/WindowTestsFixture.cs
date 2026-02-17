// <copyright file="WindowTestsFixture.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Shared fixture for WindowCommandsTests.
/// Each test creates its own file via CreateTestFile().
/// </summary>
public class WindowTestsFixture : IAsyncLifetime
{
    private readonly string _tempDir;

    public WindowTestsFixture()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"WindowTests_{Guid.NewGuid():N}");
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

    public Task InitializeAsync() => Task.CompletedTask;

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
