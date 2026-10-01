using Sbroenne.ExcelMcp.Generated;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal enum MacCapabilityTier
{
    Native,
    OfficeAddIn,
    OptionalNativeHelper,
    Unsupported
}

internal sealed record MacCommandCapability(
    string Command,
    bool IsAvailable,
    MacCapabilityTier RequiredTier,
    string UnavailableMessage,
    string WindowsSemantics,
    IReadOnlyList<string> WindowsVariants,
    string ImplementationStatus,
    string Evidence,
    string ExcelApiVersion,
    string Blocker,
    string PlannedTier,
    string AcceptanceFixture,
    string AcceptanceCommand,
    string RecoveryRule,
    string EvidenceCriteria);

internal static class MacCommandCapabilities
{
    private static readonly IReadOnlyList<MacCommandCapability> All =
        MacActionInventory.Actions.Select(ToCapability).ToArray();

    private static readonly Dictionary<string, MacCommandCapability> ByCommand =
        All.ToDictionary(item => item.Command, StringComparer.Ordinal);

    public static IReadOnlyList<MacCommandCapability> Inventory => All;

    public static string InventoryJson => MacActionInventory.Json;

    public static MacCommandCapability Get(
        string command,
        bool officeCandidateEnabled = false)
    {
        if (ByCommand.TryGetValue(command, out var capability))
        {
            if (officeCandidateEnabled && MacOfficeActionCatalog.TryGet(command, out _))
            {
                return capability with
                {
                    IsAvailable = true,
                    RequiredTier = MacCapabilityTier.OfficeAddIn,
                    UnavailableMessage = string.Empty
                };
            }
            return capability;
        }

        return new MacCommandCapability(
            command,
            false,
            MacCapabilityTier.Unsupported,
            $"Command '{command}' requires a capability that is not supported by the macOS Excel backend. It remains available on Windows.",
            "No generated Windows contract exists for this command.",
            [],
            "NotTested",
            "The command is absent from the generated public action inventory.",
            "Unverified.",
            "a capability that is not supported by the macOS Excel backend",
            "MacLimitationCandidate",
            "No generated acceptance fixture exists.",
            "Add the command to the generated inventory before acceptance.",
            "Do not mutate an unknown command.",
            "A generated contract and exact CLI/MCP evidence are required.");
    }

    private static MacCommandCapability ToCapability(MacActionInventoryItem item)
    {
        if (!Enum.TryParse<MacCapabilityTier>(item.Tier, out var tier))
        {
            throw new InvalidOperationException(
                $"Generated macOS capability tier '{item.Tier}' for '{item.Command}' is invalid.");
        }

        return new MacCommandCapability(
            item.Command,
            item.IsAvailable,
            tier,
            item.IsAvailable
                ? string.Empty
                : $"Command '{item.Command}' is unavailable on macOS: " +
                  $"{item.Blocker.TrimEnd('.')}. It remains available on Windows.",
            item.WindowsSemantics,
            item.WindowsVariants,
            item.ImplementationStatus,
            item.Evidence,
            item.ExcelApiVersion,
            item.Blocker,
            item.PlannedTier,
            item.AcceptanceFixture,
            item.AcceptanceCommand,
            item.RecoveryRule,
            item.EvidenceCriteria);
    }
}
