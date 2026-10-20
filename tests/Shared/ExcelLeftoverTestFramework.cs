using System.Reflection;
using Sbroenne.ExcelMcp.Tests.Shared;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace Sbroenne.ExcelMcp.Tests.Infrastructure;

public sealed class ExcelLeftoverTestFramework(IMessageSink messageSink) : XunitTestFramework(messageSink)
{
    protected override ITestFrameworkExecutor CreateExecutor(AssemblyName assemblyName) =>
        new LeftoverExecutor(assemblyName, SourceInformationProvider, DiagnosticMessageSink);

    private sealed class LeftoverExecutor(
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
            using var runner = new LeftoverAssemblyRunner(
                TestAssembly, testCases, DiagnosticMessageSink, executionMessageSink, executionOptions);
            await runner.RunAsync();
        }
    }

    private sealed class LeftoverAssemblyRunner(
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
            Aggregator.Run(TestRunExcelLeftoverCheck.Start);
        }

        protected override async Task BeforeTestAssemblyFinishedAsync()
        {
            await base.BeforeTestAssemblyFinishedAsync();
            // Report leftovers through xUnit before it announces assembly completion.
            Aggregator.Run(TestRunExcelLeftoverCheck.Finish);
        }
    }
}
