// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Text.Json;

namespace Sbroenne.ExcelMcp.Service.Safety;

internal sealed record ReviewAuthorization(
    string ReviewId,
    string OperationId,
    string SessionId,
    string Command,
    string NormalizedArgs,
    string WorkbookIdentity,
    bool CheckpointRequested,
    CheckpointReservation? CheckpointReservation,
    string BaselineFingerprint,
    DateTime ReviewedAtUtc,
    DateTime ExpiresAtUtc,
    SafetyScope Scope);

internal sealed record SafetyScope(
    IReadOnlyList<string> Sheets,
    IReadOnlyList<string> Ranges,
    IReadOnlyList<string> Objects)
{
    public static SafetyScope Workbook { get; } = new([], [], ["workbook"]);
}

internal sealed record SafetyScopeSummary(
    int SheetCount,
    int RangeCount,
    int ObjectCount)
{
    public static SafetyScopeSummary From(SafetyScope scope) => new(
        scope.Sheets.Count,
        scope.Ranges.Count,
        scope.Objects.Count);
}

internal sealed record SafetyArgumentSummary(
    int ParameterCount,
    int StringCount,
    int NumberCount,
    int BooleanCount,
    int ObjectCount,
    int ArrayCount,
    int NullCount)
{
    public static SafetyArgumentSummary Empty { get; } = new(0, 0, 0, 0, 0, 0, 0);

    public static SafetyArgumentSummary FromJson(string? argsJson)
    {
        if (string.IsNullOrWhiteSpace(argsJson))
        {
            return Empty;
        }

        try
        {
            using var document = JsonDocument.Parse(argsJson);
            if (document.RootElement.ValueKind != JsonValueKind.Object)
            {
                return Empty;
            }

            var parameterCount = 0;
            var stringCount = 0;
            var numberCount = 0;
            var booleanCount = 0;
            var objectCount = 0;
            var arrayCount = 0;
            var nullCount = 0;

            foreach (var property in document.RootElement.EnumerateObject())
            {
                parameterCount++;
                switch (property.Value.ValueKind)
                {
                    case JsonValueKind.String:
                        stringCount++;
                        break;
                    case JsonValueKind.Number:
                        numberCount++;
                        break;
                    case JsonValueKind.True:
                    case JsonValueKind.False:
                        booleanCount++;
                        break;
                    case JsonValueKind.Object:
                        objectCount++;
                        break;
                    case JsonValueKind.Array:
                        arrayCount++;
                        break;
                    case JsonValueKind.Null:
                    case JsonValueKind.Undefined:
                        nullCount++;
                        break;
                }
            }

            return new SafetyArgumentSummary(
                parameterCount,
                stringCount,
                numberCount,
                booleanCount,
                objectCount,
                arrayCount,
                nullCount);
        }
        catch (JsonException)
        {
            return Empty;
        }
    }
}

// Fingerprint authorizes the target plus workbook structure. VerificationFingerprint
// is target-comparable and is the only fingerprint domain exposed in receipts.
internal sealed record SemanticSnapshot(
    string Fingerprint,
    string VerificationFingerprint,
    SafetyScope Scope,
    string VerificationLevel,
    bool IsBounded,
    IReadOnlyList<string> CellHashes,
    int CellCount,
    int RowCount = 0,
    int ColumnCount = 0);

internal sealed record VerificationReceipt(
    string Status,
    SafetyScope Scope,
    int ChangedCells,
    string BeforeFingerprint,
    string AfterFingerprint,
    string? Limitation);

internal sealed record VerificationSummary(
    string Status,
    SafetyScopeSummary Scope,
    int ChangedCells,
    string? Limitation)
{
    public static VerificationSummary From(VerificationReceipt receipt) => new(
        receipt.Status,
        SafetyScopeSummary.From(receipt.Scope),
        receipt.ChangedCells,
        receipt.Limitation);
}

internal sealed record SafetyTransition(
    string State,
    DateTime AtUtc,
    string? Category = null);

internal sealed record SafetyCheckpointRecord(
    string RecoveryId,
    string RelativePath,
    string Sha256,
    long Size,
    bool CalculationSettled,
    DateTime CreatedAtUtc);

internal sealed class SafetyOperationRecord
{
    public required string OperationId { get; init; }
    public required string SessionId { get; init; }
    public required string Command { get; init; }
    public required string MutationKind { get; init; }
    public required string WorkbookIdentity { get; init; }
    public required SafetyScopeSummary Affected { get; init; }
    public SafetyArgumentSummary ArgumentSummary { get; init; } = SafetyArgumentSummary.Empty;
    public required DateTime CreatedAtUtc { get; init; }
    public List<SafetyTransition> Transitions { get; init; } = [];
    public SafetyCheckpointRecord? Checkpoint { get; set; }
    public VerificationSummary? Verification { get; set; }
    public string? OutcomeCategory { get; set; }
    public long? DurationMilliseconds { get; set; }
}

internal sealed record CheckpointCreationResult(
    bool Created,
    string RecoveryId,
    string Path,
    string RelativePath,
    string Sha256,
    long Size,
    bool CalculationSettled,
    DateTime CreatedAtUtc);

internal sealed record CheckpointReservation(
    string RecoveryId,
    string AbsolutePath,
    string RelativePath);

/// <summary>
/// Durable evidence that a staged checkpoint was fully flushed before publication.
/// </summary>
internal sealed record CheckpointReadyMarker(long Size, string Sha256);

internal static class JsonResult
{
    public static JsonElement? Parse(string? json)
    {
        if (string.IsNullOrWhiteSpace(json))
        {
            return null;
        }

        try
        {
            using var document = JsonDocument.Parse(json);
            return document.RootElement.Clone();
        }
        catch (JsonException)
        {
            return null;
        }
    }
}
