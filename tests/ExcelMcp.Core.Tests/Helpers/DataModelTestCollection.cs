using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Collection definition for tests that share the DataModelPivotTableFixture.
/// This creates ONE fixture instance shared across ALL test classes in this collection.
/// 
/// Usage: Add [Collection("DataModel")] attribute to test classes that need the fixture.
/// The fixture is injected via constructor parameter.
/// 
/// Benefits:
/// - Fixture created ONCE for all test classes in the collection (~1.5 min setup)
/// - Instead of once per test class (6 classes × ~1.5 min = ~9 min setup)
/// - Saves ~7.5 minutes of test execution time
/// </summary>
[CollectionDefinition("DataModel")]
public class DataModelTestsDefinition : ICollectionFixture<DataModelPivotTableFixture>
{
    // This class has no code - it's just a marker for xUnit
    // to associate the collection name with the fixture type
}
