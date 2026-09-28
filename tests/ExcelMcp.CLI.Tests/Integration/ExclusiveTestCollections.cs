using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[CollectionDefinition("Sequential", DisableParallelization = true)]
public sealed class SequentialCollectionDefinition;

[CollectionDefinition("ConsoleOutput", DisableParallelization = true)]
public sealed class ConsoleOutputCollectionDefinition;
