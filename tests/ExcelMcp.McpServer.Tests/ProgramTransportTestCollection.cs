// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests;

/// <summary>
/// Serializes production-host tests that may own Excel processes or mutate process-wide diagnostics.
/// </summary>
/// <remarks>
/// Transports and bridges are host-owned, but Excel-dependent tests must not overlap.
/// </remarks>
[CollectionDefinition("ProgramTransport", DisableParallelization = true)]
#pragma warning disable CA1711 // xUnit collection definition requires class name ending in 'Collection' by convention
public class ProgramTransportTestCollection
#pragma warning restore CA1711
{
    // This class has no code - it's a marker for xUnit collection definition
}


