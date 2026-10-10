using System.Globalization;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

public partial class VbaCommands
{
    private const int MaximumVbaReadLines = 500;
    private const int VbPropertyLet = 1;
    private const int VbPropertySet = 2;
    private const int VbPropertyGet = 3;

    /// <inheritdoc />
    public VbaReadResult Read(
        IExcelBatch batch,
        string moduleName,
        string? procedureName,
        string? procedureKind,
        int? startLine,
        int? lineCount)
    {
        EnsureVbaFile(batch.WorkbookPath);
        ValidateReadSelection(moduleName, procedureName, procedureKind, startLine, lineCount);

        if (!IsVbaTrustEnabled())
        {
            throw new OperationFailureException(OperationFailureCategory.Permissions, VbaTrustErrorMessage);
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic? vbaProject = null;
            dynamic? vbComponents = null;
            dynamic? component = null;
            dynamic? codeModule = null;
            try
            {
                // PIA gap: VBProject is in Microsoft.Vbe.Interop, which is unavailable on supported Click-to-Run Office installs.
                vbaProject = ((dynamic)ctx.Book).VBProject;
                EnsureProjectUnlocked((object)vbaProject);
                vbComponents = vbaProject.VBComponents;
                component = FindComponent(vbComponents, moduleName);
                codeModule = component.CodeModule;

                int moduleLineCount = Convert.ToInt32(codeModule.CountOfLines, CultureInfo.InvariantCulture);
                int sourceStartLine;
                int sourceLineCount;
                string source;
                string? sourceHash;

                if (!string.IsNullOrWhiteSpace(procedureName))
                {
                    VbaProcedureInfo procedure = FindProcedure(
                        codeModule,
                        procedureName,
                        procedureKind,
                        requireKind: false,
                        cancellationToken: ct);
                    sourceStartLine = procedure.StartLine;
                    sourceLineCount = procedure.LineCount;
                    source = ReadLines(codeModule, sourceStartLine, sourceLineCount);
                    sourceHash = VbaProcedureSource.ComputeHash(source);
                }
                else
                {
                    sourceStartLine = startLine!.Value;
                    if (sourceStartLine > moduleLineCount)
                    {
                        throw new OperationFailureException(
                            OperationFailureCategory.InvalidInput,
                            $"Start line {sourceStartLine} is past the end of module '{moduleName}', which has {moduleLineCount} lines.");
                    }

                    sourceLineCount = Math.Min(lineCount!.Value, moduleLineCount - sourceStartLine + 1);
                    source = ReadLines(codeModule, sourceStartLine, sourceLineCount);
                    sourceHash = VbaProcedureSource.ComputeHash(source);
                }

                int returnedLineCount = Math.Min(sourceLineCount, MaximumVbaReadLines);
                string returnedCode = returnedLineCount == sourceLineCount
                    ? source
                    : ReadLines(codeModule, sourceStartLine, returnedLineCount);

                return new VbaReadResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    ModuleName = moduleName,
                    StartLine = sourceStartLine,
                    TotalLineCount = sourceLineCount,
                    ReturnedLineCount = returnedLineCount,
                    Code = returnedCode,
                    HasMore = returnedLineCount < sourceLineCount,
                    NextStartLine = returnedLineCount < sourceLineCount
                        ? sourceStartLine + returnedLineCount
                        : null,
                    SourceHash = sourceHash
                };
            }
            catch (COMException comEx) when (IsVbaTrustError(comEx))
            {
                throw new OperationFailureException(OperationFailureCategory.Permissions, VbaTrustErrorMessage, comEx);
            }
            catch (COMException comEx) when (comEx.ErrorCode == GenericOfficeAutomationError)
            {
                throw new InvalidOperationException(BuildGenericComErrorMessage(comEx), comEx);
            }
            finally
            {
                ComUtilities.Release(ref codeModule);
                ComUtilities.Release(ref component);
                ComUtilities.Release(ref vbComponents);
                ComUtilities.Release(ref vbaProject);
            }
        });
    }

    /// <inheritdoc />
    public VbaProcedureEditResult ReplaceProcedure(
        IExcelBatch batch,
        string moduleName,
        string procedureName,
        string procedureKind,
        string expectedSourceHash,
        string vbaCode)
    {
        EnsureVbaFile(batch.WorkbookPath);
        ValidateProcedureReplacement(moduleName, procedureName, procedureKind, expectedSourceHash, vbaCode);

        if (!IsVbaTrustEnabled())
        {
            throw new OperationFailureException(OperationFailureCategory.Permissions, VbaTrustErrorMessage);
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic? vbaProject = null;
            dynamic? vbComponents = null;
            dynamic? component = null;
            dynamic? codeModule = null;
            string? partialState = null;
            try
            {
                // PIA gap: VBProject is in Microsoft.Vbe.Interop, which is unavailable on supported Click-to-Run Office installs.
                vbaProject = ((dynamic)ctx.Book).VBProject;
                EnsureProjectUnlocked((object)vbaProject);
                string mode = GetProjectMode(Convert.ToInt32(vbaProject.Mode, CultureInfo.InvariantCulture));
                if (mode != "Design")
                {
                    throw new OperationFailureException(OperationFailureCategory.Conflict,
                        $"The VBA project is in {mode} mode. Stop its execution in Excel before replacing source.");
                }
                vbComponents = vbaProject.VBComponents;
                component = FindComponent(vbComponents, moduleName);
                codeModule = component.CodeModule;

                VbaProcedureInfo procedure = FindProcedure(
                    codeModule,
                    procedureName,
                    procedureKind,
                    requireKind: true,
                    cancellationToken: ct);
                string currentSource = ReadLines(codeModule, procedure.StartLine, procedure.LineCount);
                string currentHash = VbaProcedureSource.ComputeHash(currentSource);
                byte[] actualHash = Convert.FromHexString(currentHash);
                byte[] expectedHash = Convert.FromHexString(expectedSourceHash);
                if (!CryptographicOperations.FixedTimeEquals(actualHash, expectedHash))
                {
                    throw new OperationFailureException(
                        OperationFailureCategory.Conflict,
                        $"Procedure '{procedureName}' changed after it was read. Read it again before replacing it.");
                }

                string bodySource = ReadLines(codeModule, procedure.BodyStartLine,
                    procedure.LineCount - (procedure.BodyStartLine - procedure.StartLine));
                int bodyLineCount = VbaProcedureSource.GetBodyLineCount(bodySource, procedure.Kind);
                codeModule.DeleteLines(procedure.BodyStartLine, bodyLineCount);
                partialState = "The selected procedure body was deleted, but its replacement has not been confirmed.";
                codeModule.InsertLines(procedure.BodyStartLine, vbaCode.TrimEnd('\r', '\n'));
                partialState = "Replacement source was inserted, but read-back verification did not complete.";

                VbaProcedureInfo updatedProcedure = FindProcedure(
                    codeModule,
                    procedureName,
                    procedureKind,
                    requireKind: true,
                    cancellationToken: ct);
                string storedSource = ReadLines(codeModule, updatedProcedure.StartLine, updatedProcedure.LineCount);

                return new VbaProcedureEditResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    Action = "replace-procedure",
                    Message = "The procedure source was replaced and read back. This does not confirm that it compiles or runs correctly.",
                    ProcedureName = updatedProcedure.Name,
                    ProcedureKind = updatedProcedure.Kind,
                    SourceHash = VbaProcedureSource.ComputeHash(storedSource),
                    StartLine = updatedProcedure.StartLine,
                    LineCount = updatedProcedure.LineCount
                };
            }
            catch (COMException comEx) when (partialState is not null)
            {
                throw new InvalidOperationException(
                    $"{partialState} Module '{moduleName}', procedure '{procedureName}': {comEx.Message} No rollback was attempted. Inspect the module before continuing.",
                    comEx);
            }
            catch (OperationFailureException ex) when (partialState is not null)
            {
                throw new OperationFailureException(ex.ErrorCategory,
                    $"{partialState} Module '{moduleName}', procedure '{procedureName}': {ex.Message} No rollback was attempted. Inspect the module before continuing.",
                    ex);
            }
            catch (COMException comEx) when (IsVbaTrustError(comEx))
            {
                throw new OperationFailureException(OperationFailureCategory.Permissions, VbaTrustErrorMessage, comEx);
            }
            catch (COMException comEx) when (comEx.ErrorCode == GenericOfficeAutomationError)
            {
                throw new InvalidOperationException(BuildGenericComErrorMessage(comEx), comEx);
            }
            finally
            {
                ComUtilities.Release(ref codeModule);
                ComUtilities.Release(ref component);
                ComUtilities.Release(ref vbComponents);
                ComUtilities.Release(ref vbaProject);
            }
        });
    }

    internal static List<VbaProcedureInfo> GetProcedures(object codeModuleObject, CancellationToken cancellationToken = default)
    {
        dynamic codeModule = codeModuleObject;
        int moduleLineCount = Convert.ToInt32(codeModule.CountOfLines, CultureInfo.InvariantCulture);
        var procedures = new List<VbaProcedureInfo>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (int line = 1; line <= moduleLineCount; line++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string firstLine = Convert.ToString(codeModule.Lines[line, 1], CultureInfo.InvariantCulture) ?? string.Empty;
            VbaProcedureDeclaration? declaration = VbaProcedureSource.TryParseProcedureHeader(firstLine);
            if (declaration == null)
            {
                continue;
            }

            int procedureKind = GetVbProcedureKind(declaration.Kind);
            int startLine = Convert.ToInt32(
                codeModule.ProcStartLine(declaration.Name, procedureKind),
                CultureInfo.InvariantCulture);

            string key = $"{startLine}:{declaration.Kind}";
            if (!seen.Add(key))
            {
                continue;
            }

            int count = Convert.ToInt32(
                codeModule.ProcCountLines(declaration.Name, procedureKind),
                CultureInfo.InvariantCulture);
            procedures.Add(new VbaProcedureInfo
            {
                Name = declaration.Name,
                Kind = declaration.Kind,
                StartLine = startLine,
                BodyStartLine = Convert.ToInt32(
                    codeModule.ProcBodyLine(declaration.Name, procedureKind), CultureInfo.InvariantCulture),
                LineCount = count
            });
        }

        return procedures;
    }

    private static void ValidateReadSelection(
        string moduleName,
        string? procedureName,
        string? procedureKind,
        int? startLine,
        int? lineCount)
    {
        if (string.IsNullOrWhiteSpace(moduleName))
        {
            throw InvalidVbaInput("Module name cannot be empty.");
        }

        bool selectingProcedure = !string.IsNullOrWhiteSpace(procedureName);
        bool selectingLines = startLine.HasValue;
        if (selectingProcedure == selectingLines)
        {
            throw InvalidVbaInput("Select either a procedure name or a starting line and line count.");
        }

        if (selectingProcedure)
        {
            if (lineCount.HasValue || procedureKind is not null && VbaProcedureSource.NormalizeKind(procedureKind).Length == 0)
            {
                throw InvalidVbaInput("Procedure reads cannot include a line count, and procedureKind must be Sub, Function, Property Get, Property Let, or Property Set.");
            }
        }
        else if (procedureKind is not null || !lineCount.HasValue || lineCount.Value < 1 || lineCount.Value > MaximumVbaReadLines)
        {
            throw InvalidVbaInput($"Line reads require lineCount from 1 through {MaximumVbaReadLines}, and cannot include procedureKind.");
        }

        if (startLine.HasValue && startLine.Value < 1)
        {
            throw InvalidVbaInput("startLine must be greater than zero.");
        }
    }

    private static void ValidateProcedureReplacement(
        string moduleName,
        string procedureName,
        string procedureKind,
        string expectedSourceHash,
        string vbaCode)
    {
        if (string.IsNullOrWhiteSpace(moduleName) || string.IsNullOrWhiteSpace(procedureName))
        {
            throw InvalidVbaInput("Module and procedure names cannot be empty.");
        }

        string normalizedKind = VbaProcedureSource.NormalizeKind(procedureKind);
        if (normalizedKind.Length == 0)
        {
            throw InvalidVbaInput("procedureKind must be Sub, Function, Property Get, Property Let, or Property Set.");
        }

        if (expectedSourceHash is null || expectedSourceHash.Length != 64 ||
            expectedSourceHash.Any(character => !Uri.IsHexDigit(character)))
        {
            throw InvalidVbaInput("expectedSourceHash must be the 64-character SourceHash returned by vba.read.");
        }

        VbaProcedureDeclaration replacement;
        try
        {
            replacement = VbaProcedureSource.ParseSingleProcedure(vbaCode);
        }
        catch (ArgumentException ex)
        {
            throw new OperationFailureException(
                OperationFailureCategory.InvalidInput,
                $"Replacement code must contain exactly one complete procedure: {ex.Message}",
                ex);
        }

        if (!replacement.Name.Equals(procedureName, StringComparison.OrdinalIgnoreCase) ||
            !replacement.Kind.Equals(normalizedKind, StringComparison.OrdinalIgnoreCase))
        {
            throw InvalidVbaInput("Replacement code must have the same procedure name and kind as the selected procedure.");
        }
    }

    private static VbaProcedureInfo FindProcedure(
        dynamic codeModule,
        string procedureName,
        string? procedureKind,
        bool requireKind,
        CancellationToken cancellationToken)
    {
        string normalizedKind = VbaProcedureSource.NormalizeKind(procedureKind ?? string.Empty);
        List<VbaProcedureInfo> matches = GetProcedures((object)codeModule, cancellationToken)
            .Where(procedure => procedure.Name.Equals(procedureName, StringComparison.OrdinalIgnoreCase))
            .Where(procedure => normalizedKind.Length == 0 ||
                procedure.Kind.Equals(normalizedKind, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (matches.Count == 0)
        {
            throw new OperationFailureException(
                OperationFailureCategory.NotFound,
                $"Procedure '{procedureName}' was not found in the selected module.");
        }

        if (matches.Count > 1 || requireKind && normalizedKind.Length == 0)
        {
            throw InvalidVbaInput($"Procedure '{procedureName}' is ambiguous. Supply its procedureKind.");
        }

        return matches[0];
    }

    private static dynamic FindComponent(dynamic vbComponents, string moduleName)
    {
        int componentCount = Convert.ToInt32(vbComponents.Count, CultureInfo.InvariantCulture);
        for (int index = 1; index <= componentCount; index++)
        {
            dynamic? component = null;
            try
            {
                component = vbComponents.Item(index);
                string name = Convert.ToString(component.Name, CultureInfo.InvariantCulture) ?? string.Empty;
                if (name.Equals(moduleName, StringComparison.OrdinalIgnoreCase))
                {
                    dynamic target = component;
                    component = null;
                    return target;
                }
            }
            finally
            {
                ComUtilities.Release(ref component);
            }
        }

        throw new OperationFailureException(OperationFailureCategory.NotFound, $"Module '{moduleName}' not found.");
    }

    private static string ReadLines(dynamic codeModule, int startLine, int count)
    {
        return count <= 0
            ? string.Empty
            : Convert.ToString(codeModule.Lines[startLine, count], CultureInfo.InvariantCulture) ?? string.Empty;
    }

    private static int GetVbProcedureKind(string kind)
    {
        return kind switch
        {
            "Property Let" => VbPropertyLet,
            "Property Set" => VbPropertySet,
            "Property Get" => VbPropertyGet,
            _ => 0
        };
    }

    private static OperationFailureException InvalidVbaInput(string message)
    {
        return new OperationFailureException(OperationFailureCategory.InvalidInput, message);
    }
}
