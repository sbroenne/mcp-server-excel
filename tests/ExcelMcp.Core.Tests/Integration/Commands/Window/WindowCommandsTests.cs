// <copyright file="WindowCommandsTests.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using Sbroenne.ExcelMcp.Core.Commands.Window;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.Window;

/// <summary>
/// Integration tests for Window management commands.
/// Tests visibility, window state, positioning, arrange presets, and status bar.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Window")]
public partial class WindowCommandsTests : IClassFixture<WindowTestsFixture>
{
    private readonly WindowCommands _commands;
    private readonly WindowTestsFixture _fixture;

    public WindowCommandsTests(WindowTestsFixture fixture)
    {
        _commands = new WindowCommands();
        _fixture = fixture;
    }
}
