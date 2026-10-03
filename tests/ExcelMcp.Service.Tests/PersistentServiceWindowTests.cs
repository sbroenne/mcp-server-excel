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

        try
        {
            await _fixture.DisposeAsync();
        }
        catch (Exception ex)
        {
            resetFailure = PersistentServiceCleanupFailures.Combine(resetFailure, ex);
        }
        if (resetFailure is not null)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw(resetFailure);
        }
    }

    private void ResetWindow()
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_commands.ClearStatusBar(batch));
        Assert.False(Assert.IsType<bool>(
            _fixture.ExecuteRawVerification((context, _) => context.App.StatusBar)));
        RequireSuccess(_commands.SetState(batch, "normal"));
        RequireSuccess(_commands.Hide(batch));
    }

    protected static T RequireSuccess<T>(T result) where T : Sbroenne.ExcelMcp.Core.Models.ResultBase
    {
        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage), result.ErrorMessage);
        return result;
    }
}
