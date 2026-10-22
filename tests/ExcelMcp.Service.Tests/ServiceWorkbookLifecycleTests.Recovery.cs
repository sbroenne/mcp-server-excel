using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Fact]
    public async Task AttachedOperationCancellation_PreservesErrorAndRevealsUntrackedExcel()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownerId = await CreateSessionAsync(service, Path.Join(directory, "attached-recovery-owner.xlsx"));
            sessions[ownerId] = 0;
            var owner = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownerId));
            var workbookName = owner.Execute((context, _) => AddReviewWorkbook(context.App, "cancelled attached edit"));
            var attachedId = await AttachReviewWorkbookAsync(
                service,
                await DiscoverReviewWorkbookAsync(service, workbookName));
            sessions[attachedId] = 0;
            var attached = Assert.IsType<AttachedWorkbookBatch>(
                service.SessionManager.GetSession(attachedId));
            var identity = Assert.IsType<ExcelProcessIdentity>(owner.ProcessIdentity);
            Assert.Contains(identity, SessionManager.GetTrackedExcelProcesses());
            SessionManager.UntrackExcelProcess(identity);
            Assert.DoesNotContain(identity, SessionManager.GetTrackedExcelProcesses());

            var notifications = 0;
            Action<ExcelProcessIdentity> rejectHandoverNotification = _ =>
            {
                Interlocked.Increment(ref notifications);
                throw new InvalidOperationException("Recovery must not be reported as a completed handover.");
            };
            SessionManager.ExcelProcessIdentityPreserved += rejectHandoverNotification;
            using var entered = new ManualResetEventSlim();
            using var release = new ManualResetEventSlim();
            using var cancellation = new CancellationTokenSource();
            Task? pending = null;
            try
            {
                RequireSuccess(await service.ProcessAsync(new ServiceRequest
                {
                    Command = "window.hide",
                    SessionId = attachedId
                }));
                Assert.False(owner.IsExcelVisible);

                pending = Task.Run(() => attached.Execute((_, _) =>
                {
                    entered.Set();
                    if (!release.Wait(TimeSpan.FromSeconds(10)))
                        throw new TimeoutException("Test did not release the attached operation.");
                }, cancellation.Token));
                Assert.True(entered.Wait(TimeSpan.FromSeconds(5)));
                cancellation.Cancel();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await pending);
                release.Set();
                Assert.True(SpinWait.SpinUntil(
                    () => attached.GetRefreshState() == WorkbookRefreshState.Ready,
                    TimeSpan.FromSeconds(10)));
                Assert.Equal(0, notifications);
                Assert.False(SessionManager.IsExcelProcessPreserved(identity));
                Assert.True(owner.IsExcelVisible);
                Assert.Equal(ownerId, Assert.Single(service.SessionManager.GetActiveSessions(),
                    session => session.SessionId == ownerId).SessionId);
                Assert.Equal(attachedId, Assert.Single(service.SessionManager.GetActiveSessions(),
                    session => session.SessionId == attachedId).SessionId);
            }
            finally
            {
                release.Set();
                SessionManager.ExcelProcessIdentityPreserved -= rejectHandoverNotification;
                if (pending is not null)
                {
                    try
                    {
                        await pending;
                    }
                    catch (OperationCanceledException)
                    {
                    }
                }
                if (owner.IsExcelProcessAlive())
                    SessionManager.TrackExcelProcess(identity);
            }
        });
    }
}
