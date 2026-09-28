using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

[CollectionDefinition("PowerQuery", DisableParallelization = true)]
public sealed class PowerQueryCollectionDefinition : ICollectionFixture<PowerQueryTestsFixture>;
