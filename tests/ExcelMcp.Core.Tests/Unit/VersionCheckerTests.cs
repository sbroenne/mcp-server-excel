using Sbroenne.ExcelMcp.Core;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

/// <summary>
/// Unit tests for VersionChecker functionality
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
public class VersionCheckerTests
{
    [Fact]
    public void GetCurrentVersion_ShouldReturnValidVersion()
    {
        // Act
        var version = VersionChecker.GetCurrentVersion();

        // Assert
        Assert.NotNull(version);
        Assert.True(version.Major >= 1);
    }

    [Fact]
    public async Task CheckForUpdatesAsync_WithInvalidPackageId_ShouldReturnFailure()
    {
        // Arrange
        var checker = new VersionChecker();
        var invalidPackageId = "NonExistent.Package.That.Does.Not.Exist.12345";

        // Act
        var result = await checker.CheckForUpdatesAsync(invalidPackageId);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task CheckForUpdatesAsync_WithValidPackageId_ShouldReturnSuccess()
    {
        // Arrange
        var checker = new VersionChecker();
        // Use a well-known package for testing
        var packageId = "Newtonsoft.Json";

        // Act
        var result = await checker.CheckForUpdatesAsync(packageId);

        // Assert
        Assert.True(result.Success, $"Failed to check version: {result.ErrorMessage}");
        Assert.NotNull(result.LatestVersion);
        Assert.NotNull(result.PackageId);
        Assert.Equal(packageId, result.PackageId);
    }

    [Fact]
    public void VersionCheckResult_GetMessage_WhenNotSuccessful_ShouldReturnWarning()
    {
        // Arrange
        var result = new VersionCheckResult
        {
            Success = false,
            ErrorMessage = "Network error"
        };

        // Act
        var message = result.GetMessage();

        // Assert
        Assert.Contains("Warning", message);
        Assert.Contains("Network error", message);
    }

    [Fact]
    public void VersionCheckResult_GetMessage_WhenOutdated_ShouldReturnUpdateMessage()
    {
        // Arrange
        var result = new VersionCheckResult
        {
            Success = true,
            IsOutdated = true,
            CurrentVersion = "1.0.0",
            LatestVersion = "2.0.0",
            PackageId = "TestPackage"
        };

        // Act
        var message = result.GetMessage();

        // Assert
        Assert.Contains("Warning", message);
        Assert.Contains("2.0.0", message);
        Assert.Contains("1.0.0", message);
        Assert.Contains("dotnet tool update", message);
    }

    [Fact]
    public void VersionCheckResult_GetMessage_WhenUpToDate_ShouldReturnSuccessMessage()
    {
        // Arrange
        var result = new VersionCheckResult
        {
            Success = true,
            IsOutdated = false,
            CurrentVersion = "1.0.0",
            LatestVersion = "1.0.0"
        };

        // Act
        var message = result.GetMessage();

        // Assert
        Assert.Contains("latest version", message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("1.0.0", message);
    }
}
