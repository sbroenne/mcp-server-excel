using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Public, closed values for configuring opt-in workbook safety behaviour.
/// </summary>
public sealed class SafetyConfigurationOptions
{
    /// <summary>Configures whether mutations can require a prior review.</summary>
    public SafetyReviewMode? ReviewMode { get; init; }
    /// <summary>Configures whether mutations create a prior-state checkpoint.</summary>
    public SafetyCheckpointMode? CheckpointMode { get; init; }
    /// <summary>Configures whether operation transitions are recorded durably.</summary>
    public SafetyJournalMode? JournalMode { get; init; }
    /// <summary>Configures whether mutations receive a semantic verification receipt.</summary>
    public SafetyVerificationMode? VerificationMode { get; init; }
    /// <summary>Configures the action taken for an abnormal safety-enabled shutdown.</summary>
    public SafetyAbnormalShutdownPolicy? AbnormalShutdownPolicy { get; init; }
}

/// <summary>Controls the review requirement for consequential workbook mutations.</summary>
public enum SafetyReviewMode
{
    /// <summary>Do not request or require a review.</summary>
    [JsonStringEnumMemberName("off")]
    Off,
    /// <summary>Allow a caller to request a review without requiring one.</summary>
    [JsonStringEnumMemberName("optional")]
    Optional,
    /// <summary>Require a valid review before executing a mutation.</summary>
    [JsonStringEnumMemberName("required")]
    Required
}

/// <summary>Controls the checkpoint requirement for consequential workbook mutations.</summary>
public enum SafetyCheckpointMode
{
    /// <summary>Do not create checkpoints automatically.</summary>
    [JsonStringEnumMemberName("off")]
    Off,
    /// <summary>Create a checkpoint when a caller explicitly requests one.</summary>
    [JsonStringEnumMemberName("onRequest")]
    OnRequest,
    /// <summary>Require a valid checkpoint before each mutation.</summary>
    [JsonStringEnumMemberName("required")]
    Required
}

/// <summary>Controls durable operation-journal recording.</summary>
public enum SafetyJournalMode
{
    /// <summary>Do not persist journal records.</summary>
    [JsonStringEnumMemberName("off")]
    Off,
    /// <summary>Persist operation journal records.</summary>
    [JsonStringEnumMemberName("on")]
    On
}

/// <summary>Controls semantic verification after workbook mutations.</summary>
public enum SafetyVerificationMode
{
    /// <summary>Do not generate verification receipts.</summary>
    [JsonStringEnumMemberName("off")]
    Off,
    /// <summary>Generate bounded semantic verification receipts.</summary>
    [JsonStringEnumMemberName("on")]
    On
}

/// <summary>Controls the session outcome when a safety-enabled shutdown is abnormal.</summary>
public enum SafetyAbnormalShutdownPolicy
{
    /// <summary>Preserve the legacy automatic-save shutdown behaviour.</summary>
    [JsonStringEnumMemberName("legacyAutoSave")]
    LegacyAutoSave,
    /// <summary>Discard active state while preserving durable recovery evidence.</summary>
    [JsonStringEnumMemberName("discardWithRecoveryEvidence")]
    DiscardWithRecoveryEvidence
}
