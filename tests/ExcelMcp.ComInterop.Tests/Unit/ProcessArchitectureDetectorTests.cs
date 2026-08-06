// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Layer", "ComInterop")]
[Trait("Category", "Unit")]
[Trait("Feature", "SessionManager")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ProcessArchitectureDetectorTests
{
    [Fact]
    public void GetBitness_CurrentProcess_ReportsActualTargetProcessBitness()
    {
        var result = ProcessArchitectureDetector.GetBitness(Environment.ProcessId);

        Assert.Equal(Environment.Is64BitProcess ? "x64" : "x86", result);
    }

    [Theory]
    [InlineData(null)]
    [InlineData(-1)]
    public void GetBitness_MissingProcess_ReturnsUnknown(int? processId)
    {
        Assert.Equal("unknown", ProcessArchitectureDetector.GetBitness(processId));
    }
}
