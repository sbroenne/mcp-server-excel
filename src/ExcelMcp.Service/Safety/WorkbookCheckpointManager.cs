// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Service.Safety;

internal static class WorkbookCheckpointManager
{
    private static readonly TimeSpan CalculationSettleTimeout = TimeSpan.FromSeconds(5);

    public static CheckpointCreationResult Create(
        IExcelBatch batch,
        DurableSafetyStore store,
        CheckpointReservation? reservation = null)
    {
        if (!File.Exists(batch.WorkbookPath))
        {
            throw new InvalidOperationException(
                "No prior-state checkpoint can be created because this workbook has not been saved to disk.");
        }

        var allocation = reservation ?? store.AllocateCheckpoint(batch.WorkbookPath);
        var pendingPath = GetPendingCheckpointPath(allocation.AbsolutePath);
        var readyMarkerPath = GetReadyMarkerPath(allocation.AbsolutePath);
        try
        {
            store.EnsureSafeCheckpointPath(allocation.AbsolutePath);
            store.EnsureSafeCheckpointPath(pendingPath);
            store.EnsureSafeCheckpointPath(readyMarkerPath);
            if (File.Exists(pendingPath))
            {
                throw new IOException("The checkpoint staging path already exists.");
            }

            var calculationSettled = batch.Execute((context, cancellationToken) =>
            {
                var deadline = DateTime.UtcNow + CalculationSettleTimeout;
                var settled = IsCalculationSettled(context.App.CalculationState);
                while (!settled && DateTime.UtcNow < deadline)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    Thread.Sleep(TimeSpan.FromMilliseconds(50));
                    settled = IsCalculationSettled(context.App.CalculationState);
                }

                cancellationToken.ThrowIfCancellationRequested();
                context.Book.SaveCopyAs(pendingPath);
                return settled;
            });
            store.EnsureSafeCheckpointPath(pendingPath);

            if (!File.Exists(pendingPath))
            {
                throw new IOException("Excel did not create the requested staged checkpoint file.");
            }

            var stagedFileInfo = new FileInfo(pendingPath);
            if (stagedFileInfo.Length <= 0)
            {
                throw new IOException("Excel created an empty checkpoint file.");
            }

            DurableFileWriter.FlushExistingFile(pendingPath);
            var stagedFileInfoAfterFlush = new FileInfo(pendingPath);
            var stagedHash = DurableSafetyStore.ComputeFileHash(pendingPath);
            DurableFileWriter.WriteUtf8Atomically(
                readyMarkerPath,
                JsonSerializer.Serialize(
                    new CheckpointReadyMarker(stagedFileInfoAfterFlush.Length, stagedHash),
                    ServiceProtocol.JsonOptions));
            DurableFileWriter.PublishFlushedFileAtomically(pendingPath, allocation.AbsolutePath);
            store.EnsureSafeCheckpointPath(allocation.AbsolutePath);
            TryDeleteReadyMarker(readyMarkerPath);

            var fileInfo = new FileInfo(allocation.AbsolutePath);
            if (fileInfo.Length <= 0)
            {
                throw new IOException("The atomically published checkpoint file is empty.");
            }

            var createdAtUtc = DateTime.UtcNow;
            return new CheckpointCreationResult(
                Created: true,
                allocation.RecoveryId,
                allocation.AbsolutePath,
                allocation.RelativePath,
                DurableSafetyStore.ComputeFileHash(allocation.AbsolutePath),
                fileInfo.Length,
                calculationSettled,
                createdAtUtc);
        }
        catch
        {
            try
            {
                store.EnsureSafeCheckpointPath(allocation.AbsolutePath);
                if (File.Exists(allocation.AbsolutePath))
                {
                    File.Delete(allocation.AbsolutePath);
                }

                store.EnsureSafeCheckpointPath(pendingPath);
                if (File.Exists(pendingPath))
                {
                    File.Delete(pendingPath);
                }

                store.EnsureSafeCheckpointPath(readyMarkerPath);
                if (File.Exists(readyMarkerPath))
                {
                    File.Delete(readyMarkerPath);
                }
            }
            catch (IOException)
            {
                // Preserve the original checkpoint failure.
            }
            catch (UnauthorizedAccessException)
            {
                // Preserve the original checkpoint failure.
            }
            catch (InvalidOperationException)
            {
                // Preserve the original checkpoint failure without following an unsafe path.
            }

            throw;
        }
    }

    internal static string GetPendingCheckpointPath(string destinationPath)
    {
        var directory = Path.GetDirectoryName(destinationPath) ??
            throw new ArgumentException("Checkpoint destination must have a parent directory.", nameof(destinationPath));
        var extension = Path.GetExtension(destinationPath);
        var baseName = Path.GetFileNameWithoutExtension(destinationPath);
        return Path.Combine(directory, $".{baseName}.pending{extension}");
    }

    internal static string GetReadyMarkerPath(string destinationPath) =>
        $"{GetPendingCheckpointPath(destinationPath)}.ready";

    internal static bool TryReadReadyMarker(
        string markerPath,
        out CheckpointReadyMarker? marker)
    {
        marker = null;
        try
        {
            var json = File.ReadAllText(markerPath);
            marker = JsonSerializer.Deserialize<CheckpointReadyMarker>(
                json,
                ServiceProtocol.JsonOptions);
            return marker is not null &&
                marker.Size > 0 &&
                !string.IsNullOrWhiteSpace(marker.Sha256) &&
                marker.Sha256.Length == 64 &&
                marker.Sha256.All(static character => Uri.IsHexDigit(character));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException or NotSupportedException)
        {
            return false;
        }
    }

    private static void TryDeleteReadyMarker(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException)
        {
            // Marker cleanup is best effort after the final checkpoint is published.
        }
        catch (UnauthorizedAccessException)
        {
            // Marker cleanup is best effort after the final checkpoint is published.
        }
    }

    private static bool IsCalculationSettled(object calculationState) =>
        Convert.ToInt32(calculationState, CultureInfo.InvariantCulture) == 0;
}
