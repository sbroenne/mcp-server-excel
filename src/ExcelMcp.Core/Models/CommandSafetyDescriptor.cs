// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Generated safety metadata for one public service action.
/// </summary>
public sealed record CommandSafetyDescriptor(
    string Command,
    bool IsMutation,
    string MutationKind,
    string ScopeResolver,
    string VerificationLevel,
    bool CheckpointRecommended,
    string RecoveryRisk,
    bool ExplicitlyClassified)
{
    /// <summary>
    /// Creates a fail-closed descriptor for a command unknown to the generated catalog.
    /// </summary>
    public static CommandSafetyDescriptor Unknown(string command) => new(
        command,
        IsMutation: true,
        MutationKind: "unknown",
        ScopeResolver: "workbook",
        VerificationLevel: "notVerified",
        CheckpointRecommended: true,
        RecoveryRisk: "high",
        ExplicitlyClassified: false);
}
