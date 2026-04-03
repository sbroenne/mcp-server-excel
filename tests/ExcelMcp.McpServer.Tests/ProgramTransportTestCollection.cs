// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests;

/// <summary>
/// Collection definition for tests that use Program.ConfigureTestTransport().
/// These tests MUST run sequentially because the in-memory MCP host uses a shared static transport hook.
/// </summary>
/// <remarks>
/// Any test that uses Program.ConfigureTestTransport() or mutates ServiceBridge test state
/// must join this collection so the shared transport and in-process service lifecycle stay serialized.
/// </remarks>
[CollectionDefinition("ProgramTransport")]
#pragma warning disable CA1711 // xUnit collection definition requires class name ending in 'Collection' by convention
public class ProgramTransportTestCollection
#pragma warning restore CA1711
{
    // This class has no code - it's a marker for xUnit collection definition
}




