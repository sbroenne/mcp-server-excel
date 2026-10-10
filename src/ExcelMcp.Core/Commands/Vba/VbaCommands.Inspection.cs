using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

public partial class VbaCommands
{
    /// <inheritdoc />
    public VbaProjectStatusResult Status(IExcelBatch batch)
    {
        EnsureVbaFile(batch.WorkbookPath);
        return batch.Execute((ctx, ct) =>
        {
            dynamic? project = null;
            try
            {
                // PIA gap: the VBA editor object model is not available through the supported Excel PIA.
                project = ((dynamic)ctx.Book).VBProject;
                return new VbaProjectStatusResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    ProjectAccess = true,
                    ProjectName = Convert.ToString(project.Name, CultureInfo.InvariantCulture),
                    Protection = GetProjectProtection(Convert.ToInt32(project.Protection, CultureInfo.InvariantCulture)),
                    Mode = GetProjectMode(Convert.ToInt32(project.Mode, CultureInfo.InvariantCulture))
                };
            }
            catch (COMException ex) when (IsVbaTrustError(ex))
            {
                return new VbaProjectStatusResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    ProjectAccess = false,
                    AccessMessage = VbaTrustErrorMessage
                };
            }
            catch (COMException ex) when (ex.ErrorCode == GenericOfficeAutomationError)
            {
                throw new InvalidOperationException(BuildGenericComErrorMessage(ex), ex);
            }
            finally
            {
                ComUtilities.Release(ref project);
            }
        });
    }

    /// <inheritdoc />
    public VbaReferencesResult References(IExcelBatch batch)
    {
        EnsureVbaFile(batch.WorkbookPath);
        return batch.Execute((ctx, ct) =>
        {
            dynamic? project = null;
            dynamic? references = null;
            try
            {
                // PIA gap: the VBA editor object model is not available through the supported Excel PIA.
                project = ((dynamic)ctx.Book).VBProject;
                EnsureProjectUnlocked((object)project);
                references = project.References;
                int count = Convert.ToInt32(references.Count, CultureInfo.InvariantCulture);
                var result = new VbaReferencesResult { Success = true, FilePath = batch.WorkbookPath };
                for (int index = 1; index <= count; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    dynamic? reference = null;
                    try
                    {
                        reference = references.Item(index);
                        VbaReferenceInfo info = ReadReferenceInfo((object)reference, index);
                        result.References.Add(info);
                        result.HasBrokenReferences |= info.IsBroken;
                    }
                    finally
                    {
                        ComUtilities.Release(ref reference);
                    }
                }

                return result;
            }
            catch (COMException ex) when (IsVbaTrustError(ex))
            {
                throw new OperationFailureException(OperationFailureCategory.Permissions, VbaTrustErrorMessage, ex);
            }
            catch (COMException ex) when (ex.ErrorCode == GenericOfficeAutomationError)
            {
                throw new InvalidOperationException(BuildGenericComErrorMessage(ex), ex);
            }
            finally
            {
                ComUtilities.Release(ref references);
                ComUtilities.Release(ref project);
            }
        });
    }

    /// <inheritdoc />
    public VbaSearchResult Search(
        IExcelBatch batch,
        string searchText,
        string? moduleName = null,
        bool wholeWord = false,
        bool matchCase = false,
        int maxMatches = 50)
    {
        EnsureVbaFile(batch.WorkbookPath);
        if (string.IsNullOrWhiteSpace(searchText) || searchText.Contains('\r') || searchText.Contains('\n'))
        {
            throw InvalidVbaInput("searchText must contain nonempty, single-line literal text.");
        }
        if (maxMatches is < 1 or > 100)
        {
            throw InvalidVbaInput("maxMatches must be from 1 through 100.");
        }
        if (moduleName is not null && string.IsNullOrWhiteSpace(moduleName))
        {
            throw InvalidVbaInput("moduleName cannot be empty when supplied.");
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic? project = null;
            dynamic? components = null;
            try
            {
                // PIA gap: the VBA editor object model is not available through the supported Excel PIA.
                project = ((dynamic)ctx.Book).VBProject;
                EnsureProjectUnlocked((object)project);
                components = project.VBComponents;
                int count = Convert.ToInt32(components.Count, CultureInfo.InvariantCulture);
                bool foundModule = moduleName is null;
                var result = new VbaSearchResult { Success = true, FilePath = batch.WorkbookPath };
                for (int index = 1; index <= count; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    dynamic? component = null;
                    dynamic? codeModule = null;
                    try
                    {
                        component = components.Item(index);
                        string name = Convert.ToString(component.Name, CultureInfo.InvariantCulture) ?? string.Empty;
                        if (moduleName is not null && !name.Equals(moduleName, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        foundModule = true;
                        codeModule = component.CodeModule;
                        if (codeModule is null)
                        {
                            continue;
                        }

                        int lineCount = Convert.ToInt32(codeModule.CountOfLines, CultureInfo.InvariantCulture);
                        int line = 1;
                        int column = 1;
                        while (line <= lineCount)
                        {
                            ct.ThrowIfCancellationRequested();
                            int endLine = lineCount;
                            int endColumn = -1;
                            bool found = Convert.ToBoolean(
                                codeModule.Find(searchText, ref line, ref column, ref endLine, ref endColumn,
                                    wholeWord, matchCase, false), CultureInfo.InvariantCulture);
                            if (!found)
                            {
                                break;
                            }
                            if (result.Matches.Count == maxMatches)
                            {
                                result.HasMore = true;
                                return result;
                            }

                            string sourceLine = ReadLines(codeModule, line, 1);
                            int excerptStart = Math.Max(0, column - 1 - 60);
                            result.Matches.Add(new VbaSearchMatch
                            {
                                ModuleName = name,
                                Line = line,
                                Column = column,
                                Excerpt = sourceLine.Substring(excerptStart, Math.Min(200, sourceLine.Length - excerptStart))
                            });

                            // Find overwrites all four bounds. Advance beyond this match and reset the end bounds above.
                            column = Math.Max(column + searchText.Length, endColumn);
                            line = endLine;
                            if (column > sourceLine.Length)
                            {
                                line++;
                                column = 1;
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref codeModule);
                        ComUtilities.Release(ref component);
                    }
                }

                if (!foundModule)
                {
                    throw new OperationFailureException(OperationFailureCategory.NotFound, $"Module '{moduleName}' not found.");
                }
                return result;
            }
            catch (COMException ex) when (IsVbaTrustError(ex))
            {
                throw new OperationFailureException(OperationFailureCategory.Permissions, VbaTrustErrorMessage, ex);
            }
            catch (COMException ex) when (ex.ErrorCode == GenericOfficeAutomationError)
            {
                throw new InvalidOperationException(BuildGenericComErrorMessage(ex), ex);
            }
            finally
            {
                ComUtilities.Release(ref components);
                ComUtilities.Release(ref project);
            }
        });
    }

    internal static VbaReferenceInfo ReadReferenceInfo(object referenceObject, int index)
    {
        dynamic reference = referenceObject;
        bool broken = Convert.ToBoolean(reference.IsBroken, CultureInfo.InvariantCulture);
        var info = new VbaReferenceInfo { Index = index, IsBroken = broken };
        // Broken references can throw when any identifying metadata is requested.
        if (!broken)
        {
            info.Name = Convert.ToString(reference.Name, CultureInfo.InvariantCulture);
            info.Description = Convert.ToString(reference.Description, CultureInfo.InvariantCulture);
            info.LibraryId = Convert.ToString(reference.Guid, CultureInfo.InvariantCulture);
            info.Major = Convert.ToInt32(reference.Major, CultureInfo.InvariantCulture);
            info.Minor = Convert.ToInt32(reference.Minor, CultureInfo.InvariantCulture);
            info.BuiltIn = Convert.ToBoolean(reference.BuiltIn, CultureInfo.InvariantCulture);
        }
        return info;
    }

    internal static string GetProjectProtection(int protection) => protection switch
    {
        0 => "None",
        1 => "Locked",
        _ => throw new InvalidOperationException($"Excel returned unknown VBA project protection value {protection}.")
    };

    internal static string GetProjectMode(int mode) => mode switch
    {
        0 => "Run",
        1 => "Break",
        2 => "Design",
        _ => throw new InvalidOperationException($"Excel returned unknown VBA project mode value {mode}.")
    };

    private static void EnsureProjectUnlocked(object projectObject)
    {
        dynamic project = projectObject;
        if (GetProjectProtection(Convert.ToInt32(project.Protection, CultureInfo.InvariantCulture)) == "Locked")
        {
            throw new OperationFailureException(OperationFailureCategory.Permissions,
                "The VBA project is password-locked. Unlock it manually in Excel before inspecting or editing its source. ExcelMcp does not unlock projects.");
        }
    }
}
