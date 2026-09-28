using System.Reflection;
using Sbroenne.ExcelMcp.Tests.Shared;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace Sbroenne.ExcelMcp.Tests.Infrastructure;

public sealed class ExcelLifetimeTestFramework(IMessageSink messageSink) : XunitTestFramework(messageSink)
{
    protected override ITestFrameworkExecutor CreateExecutor(AssemblyName assemblyName) =>
        new LifetimeExecutor(assemblyName, SourceInformationProvider, DiagnosticMessageSink);

    private sealed class LifetimeExecutor(
        AssemblyName assemblyName,
        ISourceInformationProvider sourceInformationProvider,
        IMessageSink diagnosticMessageSink)
        : XunitTestFrameworkExecutor(assemblyName, sourceInformationProvider, diagnosticMessageSink)
    {
        protected override async void RunTestCases(
            IEnumerable<IXunitTestCase> testCases,
            IMessageSink executionMessageSink,
            ITestFrameworkExecutionOptions executionOptions)
        {
            using var runner = new LifetimeAssemblyRunner(
                TestAssembly, testCases, DiagnosticMessageSink, executionMessageSink, executionOptions);
            await runner.RunAsync();
        }
    }

    private sealed class LifetimeAssemblyRunner(
        ITestAssembly testAssembly,
        IEnumerable<IXunitTestCase> testCases,
        IMessageSink diagnosticMessageSink,
        IMessageSink executionMessageSink,
        ITestFrameworkExecutionOptions executionOptions)
        : XunitTestAssemblyRunner(testAssembly, testCases, diagnosticMessageSink, executionMessageSink, executionOptions)
    {
        protected override async Task AfterTestAssemblyStartingAsync()
        {
            await base.AfterTestAssemblyStartingAsync();
            Aggregator.Run(() => TestRunExcelLifetime.StartForTestHost());
        }

        protected override async Task BeforeTestAssemblyFinishedAsync()
        {
            await base.BeforeTestAssemblyFinishedAsync();
            // VSTest can ignore ProcessExit failures after receiving completed results.
            // Report normal cleanup through xUnit before it announces assembly completion.
            Aggregator.Run(() =>
            {
                var lifetime = TestRunExcelLifetime.CurrentHost
                    ?? throw new InvalidOperationException("Test-host Excel lifetime protection was not initialized.");
                lifetime.Dispose();
            });
        }
    }
}
