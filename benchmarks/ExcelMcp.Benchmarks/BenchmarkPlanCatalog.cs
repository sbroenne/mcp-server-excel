namespace Sbroenne.ExcelMcp.Benchmarks;

internal sealed record BenchmarkPlan(
    string Id,
    string Scenario,
    string Title,
    string BaselineMeaning,
    string CandidateSuccess,
    IReadOnlyList<string> PrimaryMetrics,
    IReadOnlyList<string> ReliabilityInvariants);

internal static class BenchmarkPlanCatalog
{
    public static IReadOnlyList<BenchmarkPlan> All { get; } =
    [
        new(
            "01",
            "timeout-quarantine",
            "Quarantine timed-out sessions",
            "Measures how quickly a timed-out operation returns, whether later work fails fast, and how long owned Excel cleanup takes.",
            "A candidate improves bounded return/cleanup latency while every post-dispatch timeout remains explicitly outcome-unknown and quarantined.",
            ["return_latency_ms", "cleanup_latency_ms", "fail_fast_latency_ms", "orphan_process_count", "working_set_delta_bytes"],
            ["no_post_timeout_write", "outcome_unknown_not_success", "session_unusable_after_timeout", "no_owned_excel_orphan"]),
        new(
            "02",
            "bounded-workbook-queue",
            "Bound the per-workbook queue",
            "Measures burst admission, queue wait, memory growth, ordering, and cross-workbook isolation with the current unbounded queue.",
            "A candidate caps pending work, applies explicit backpressure or rejection without dropping mutations, and preserves FIFO workbook order.",
            ["queue_wait_ms", "burst_completion_ms", "operations_per_second", "working_set_delta_bytes", "rejected_operation_count"],
            ["fifo_order", "no_dropped_mutation", "no_duplicate_mutation", "independent_workbook_progress"]),
        new(
            "03",
            "targeted-safety-inspection",
            "Inspect only the affected workbook area",
            "Measures review and verification cost as workbook used-range and structural complexity grow under the current semantic inspector.",
            "A candidate lowers tail latency and inspected scope while manual edits, recalculation, refreshes, and structural changes still invalidate unsafe cached state.",
            ["review_latency_ms", "verification_latency_ms", "inspected_cell_count", "payload_bytes", "stale_detection_rate"],
            ["no_stale_authorization", "exact_affected_scope", "manual_edit_invalidates", "structural_change_invalidates"]),
        new(
            "04",
            "server-side-batch",
            "Execute a true server-side batch",
            "Measures N ordered service operations sent separately, including per-request safety and serialization overhead.",
            "A candidate reduces end-to-end latency, request count, payload bytes, and token estimate while preserving order and pinpointing failed operation indexes.",
            ["batch_latency_ms", "operations_per_second", "request_count", "payload_bytes", "token_estimate"],
            ["operation_order", "no_lost_operation", "no_duplicate_operation", "failure_index_reported", "session_cleanup"]),
        new(
            "05",
            "vectorized-writes",
            "Eliminate cell-by-cell COM writes",
            "Measures contiguous Range.Value2 writes and the known cell-by-cell table-append hotspot across increasing row counts.",
            "A candidate materially increases cells per second for supported rectangular writes and retains an explicit safe fallback for non-rectangular targets.",
            ["write_latency_ms", "cells_per_second", "rows_per_second", "payload_bytes", "working_set_delta_bytes"],
            ["round_trip_values_equal", "round_trip_formulas_equal", "table_shape_equal", "no_silent_multi_area_reorder"]),
        new(
            "06",
            "read-fast-path",
            "Add a fast path for ordinary reads",
            "Measures cold/warm reads, service payload size, token estimate, and refresh-to-first-consistent-read with current safety guards.",
            "A candidate lowers warm-read latency and response footprint without returning stale data after writes, calculation, refresh, VBA, or manual edits.",
            ["cold_read_ms", "warm_read_ms", "refresh_to_consistent_read_ms", "payload_bytes", "token_estimate"],
            ["round_trip_values_equal", "no_stale_read_after_write", "no_stale_read_after_refresh", "refresh_result_consistent"]),
        new(
            "07",
            "idempotency-keys",
            "Add idempotency keys",
            "Measures duplicate review execution behavior, receipt lookup cost, and semantic side effects under retries.",
            "A candidate executes a known-complete key once, returns the same receipt, rejects argument/workbook conflicts, and never auto-replays an unknown outcome.",
            ["first_execution_ms", "duplicate_retry_ms", "duplicate_execution_count", "receipt_payload_bytes", "conflict_detection_ms"],
            ["known_key_executes_once", "same_key_same_receipt", "changed_arguments_conflict", "unknown_outcome_not_replayed"]),
        new(
            "08",
            "durable-journal-checkpoints",
            "Make journals and checkpoints crash-durable",
            "Measures flushed atomic journal/checkpoint publication, transition recovery, checkpoint hashing, corruption handling, and restart visibility.",
            "A candidate publishes no malformed checkpoint, recovers every acknowledged durable transition, and keeps normal write overhead within a chosen baseline-relative budget.",
            ["journal_write_ms", "checkpoint_create_ms", "checkpoint_bytes", "restart_recovery_ms", "refresh_to_consistent_read_ms"],
            ["journal_parseable", "transition_order", "checkpoint_hash_valid", "required_checkpoint_fails_closed", "corrupt_journal_fails_closed", "no_temporary_artifacts", "recovered_state_exact"]),
        new(
            "09",
            "precise-process-tracking",
            "Track owned Excel processes precisely",
            "Measures owned-process cleanup, orphan rate, and isolation from a separately owned sentinel Excel process.",
            "A candidate validates PID, start time, and Excel identity before termination, reaches zero wrong-process kills, and cleans every owned process.",
            ["owned_process_exit_ms", "orphan_process_count", "wrong_process_kill_count", "identity_mismatch_detection_ms", "cleanup_success_rate"],
            ["sentinel_process_survives", "identity_mismatch_fails_closed", "no_wrong_process_kill", "no_owned_excel_orphan"])
    ];
}
