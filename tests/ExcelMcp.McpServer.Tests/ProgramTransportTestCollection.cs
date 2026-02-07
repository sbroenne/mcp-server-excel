// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests;

/// <summary>
/// Collection definition for tests that use Program.ConfigureTestTransport().
/// These tests MUST run sequentially because they share static state in Program.cs.
/// </summary>
/// <remarks>
/// Tests in this collection:
/// - McpServerSmokeTests
/// - ExcelFileToolOperationTrackingTests
/// 
/// Both use Program.ConfigureTestTransport() which sets static pipe fields.
/// Running them in parallel causes "writer already completed" errors.
/// </remarks>
[CollectionDefinition("ProgramTransport")]
#pragma warning disable CA1711 // xUnit collection definition requires class name ending in 'Collection' by convention
public class ProgramTransportTestCollection
#pragma warning restore CA1711
{
    // This class has no code - it's a marker for xUnit collection definition
}




