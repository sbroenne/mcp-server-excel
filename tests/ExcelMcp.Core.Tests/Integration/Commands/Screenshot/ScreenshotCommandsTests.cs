// <copyright file="ScreenshotCommandsTests.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.Screenshot;

/// <summary>
/// Integration tests for Screenshot commands.
/// Tests CaptureRange and CaptureSheet with real Excel data, charts, and tables.
/// Validates the CopyPicture retry logic that handles intermittent COM failures.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Screenshot")]
public partial class ScreenshotCommandsTests : IClassFixture<ScreenshotTestsFixture>
{
    private readonly ScreenshotCommands _commands;
    private readonly ScreenshotTestsFixture _fixture;

    public ScreenshotCommandsTests(ScreenshotTestsFixture fixture)
    {
        _commands = new ScreenshotCommands();
        _fixture = fixture;
    }
}
