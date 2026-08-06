// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

namespace Sbroenne.ExcelMcp.Service.Safety;

/// <summary>
/// Opt-in safety policy for one workbook session.
/// </summary>
internal sealed record SessionSafetyConfiguration
{
    public ReviewMode ReviewMode { get; init; } = ReviewMode.Off;
    public CheckpointMode CheckpointMode { get; init; } = CheckpointMode.Off;
    public JournalMode JournalMode { get; init; } = JournalMode.Off;
    public VerificationMode VerificationMode { get; init; } = VerificationMode.Off;
    public AbnormalShutdownPolicy AbnormalShutdownPolicy { get; init; } = AbnormalShutdownPolicy.LegacyAutoSave;

    public static SessionSafetyConfiguration Default { get; } = new();

    public bool UsesSafetyWorkflow =>
        ReviewMode != ReviewMode.Off ||
        CheckpointMode != CheckpointMode.Off ||
        JournalMode != JournalMode.Off ||
        VerificationMode != VerificationMode.Off ||
        AbnormalShutdownPolicy != AbnormalShutdownPolicy.LegacyAutoSave;

    /// <summary>
    /// Applies the safety-enabled shutdown default when callers omit the policy.
    /// Explicit policies (including the legacy auto-save policy) are preserved.
    /// </summary>
    public SessionSafetyConfiguration NormalizeAbnormalShutdownPolicy(bool abnormalShutdownPolicySpecified) =>
        !abnormalShutdownPolicySpecified && HasSafetyControls
            ? this with { AbnormalShutdownPolicy = AbnormalShutdownPolicy.DiscardWithRecoveryEvidence }
            : this;

    private bool HasSafetyControls =>
        ReviewMode != ReviewMode.Off ||
        CheckpointMode != CheckpointMode.Off ||
        JournalMode != JournalMode.Off ||
        VerificationMode != VerificationMode.Off;
}

internal enum ReviewMode
{
    Off,
    Optional,
    Required
}

internal enum CheckpointMode
{
    Off,
    OnRequest,
    Required
}

internal enum JournalMode
{
    Off,
    On
}

internal enum VerificationMode
{
    Off,
    On
}

internal enum AbnormalShutdownPolicy
{
    LegacyAutoSave,
    DiscardWithRecoveryEvidence
}
