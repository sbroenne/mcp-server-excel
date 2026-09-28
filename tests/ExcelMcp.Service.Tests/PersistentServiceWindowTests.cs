// <copyright file="WindowCommandsTests.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using Sbroenne.ExcelMcp.Core.Commands.Window;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Window")]
public sealed partial class PersistentServiceWindowTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWindowTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>;

[System.Diagnostics.CodeAnalysis.SuppressMessage(
    "Design",
    "CA1051:Do not declare visible instance fields",
    Justification = "Protected fields preserve the original tests while routing commands through Service.")]
public abstract class PersistentServiceWindowTestBase(
    PersistentServiceWorkbookFixture fixture) : IAsyncLifetime
{
    protected readonly IWindowCommands _commands =
        fixture.CreateCommands<IWindowCommands>();
    protected readonly PersistentServiceWorkbookTestScope _fixture =
        fixture.CreateScope();

    public Task InitializeAsync()
    {
        ResetWindow();
        return Task.CompletedTask;
    }

    public async Task DisposeAsync()
    {
        Exception? resetFailure = null;
        try
        {
            ResetWindow();
        }
        catch (Exception ex)
        {
            resetFailure = ex;
        }

        await _fixture.DisposeAsync();
        if (resetFailure is not null)
        {
            throw new InvalidOperationException(
                "Window state cleanup failed.",
                resetFailure);
        }
    }

    private void ResetWindow()
    {
        var batch = _fixture.BatchToken;
        _commands.ClearStatusBar(batch);
        _commands.SetState(batch, "normal");
        _commands.Hide(batch);
    }
}
