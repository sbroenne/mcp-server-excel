// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>
/// Shared pre-open validation for Excel workbook paths.
/// </summary>
public static class ExcelFileValidator
{
    /// <summary>Maximum supported workbook size (1 GiB).</summary>
    public const long MaximumFileSizeBytes = 1024L * 1024 * 1024;

    /// <summary>Maximum supported Windows path length.</summary>
    public const int MaximumPathLength = 32767;

    /// <summary>Excel SaveAs practical path limit used when creating a workbook.</summary>
    public const int MaximumCreatePathLength = 218;

    /// <summary>
    /// Inspects a workbook path without starting Excel.
    /// </summary>
    public static ExcelFileValidation Inspect(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            return ExcelFileValidation.Invalid(filePath, "File path is required.");
        }

        if (filePath.Length > MaximumPathLength)
        {
            return ExcelFileValidation.Invalid(
                filePath,
                $"File path exceeds the maximum path length of {MaximumPathLength} characters.");
        }

        string fullPath;
        try
        {
            fullPath = Path.GetFullPath(filePath);
        }
        catch (PathTooLongException)
        {
            return ExcelFileValidation.Invalid(
                filePath,
                $"File path exceeds the maximum path length of {MaximumPathLength} characters.");
        }

        if (fullPath.Length > MaximumPathLength)
        {
            return ExcelFileValidation.Invalid(
                fullPath,
                $"File path exceeds the maximum path length of {MaximumPathLength} characters.");
        }

        var extension = Path.GetExtension(fullPath).ToLowerInvariant();
        var isSupportedExtension = extension is ".xlsx" or ".xlsm";
        var isOpenableExtension = isSupportedExtension || extension == ".xls";
        var exists = File.Exists(fullPath);
        var size = exists ? new FileInfo(fullPath).Length : 0;

        var message = !exists
            ? $"File not found: {fullPath}"
            : !isSupportedExtension
                ? extension == ".xls"
                    ? "Legacy .xls workbooks can be opened, but the strict file-test policy accepts only .xlsx and .xlsm."
                    : $"Invalid file extension. Expected .xlsx or .xlsm, got {extension}"
                : size > MaximumFileSizeBytes
                    ? $"File exceeds the maximum size of {MaximumFileSizeBytes} bytes."
                    : null;

        return new ExcelFileValidation(
            fullPath,
            exists,
            size,
            extension,
            isSupportedExtension,
            isOpenableExtension,
            true,
            fullPath.Length <= MaximumCreatePathLength,
            size <= MaximumFileSizeBytes,
            message);
    }
}

/// <summary>
/// The result of pre-open workbook validation.
/// </summary>
public sealed record ExcelFileValidation(
    string FilePath,
    bool Exists,
    long Size,
    string Extension,
    bool IsSupportedExtension,
    bool IsOpenableExtension,
    bool IsWithinPathLimit,
    bool IsWithinCreatePathLimit,
    bool IsWithinSizeLimit,
    string? Message)
{
    /// <summary>Whether an existing file satisfies the strict file-test policy (.xlsx/.xlsm).</summary>
    public bool IsValidExistingWorkbook => Exists && IsSupportedExtension && IsWithinPathLimit && IsWithinSizeLimit;

    internal static ExcelFileValidation Invalid(string filePath, string message)
    {
        return new ExcelFileValidation(filePath, false, 0, string.Empty, false, false, false, false, false, message);
    }
}
