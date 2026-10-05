namespace Sbroenne.ExcelMcp.Build;

public sealed record CiCheck(string Name, string Selected, string Result);
public sealed class CiCompletionOptions
{
    public string Detection { get; set; } = "";
    public CiCheck[] Checks { get; set; } = [];
}
public static class CiCompletion
{
    public static void Verify(CiCompletionOptions options)
    {
        if (options.Detection != "success") { throw new InvalidOperationException($"Change detection did not succeed: {options.Detection}."); }
        if (options.Checks.Length != 4 || options.Checks.Select(check => check.Name).Distinct(StringComparer.Ordinal).Count() != 4 ||
            options.Checks.Any(check => check.Name is not ("tests" or "packages" or "npm" or "lockfiles")))
        {
            throw new ArgumentException("Completion requires each tests, packages, npm and lockfiles result exactly once.");
        }
        foreach (var check in options.Checks)
        {
            var expected = check.Selected switch
            {
                "true" => "success",
                "false" => "skipped",
                _ => throw new ArgumentException($"Invalid selection for {check.Name}: {check.Selected}.")
            };
            if (check.Result != expected) { throw new InvalidOperationException($"{check.Name}: expected {expected}, received {check.Result}."); }
        }
        Console.WriteLine("All selected CI checks succeeded.");
    }
}
