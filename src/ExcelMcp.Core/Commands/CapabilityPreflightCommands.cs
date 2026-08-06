using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Collects compatibility and workbook-state evidence from a live Excel session.
/// </summary>
public sealed class CapabilityPreflightCommands
{
    /// <summary>
    /// Performs a read-only capability preflight for the supplied session.
    /// </summary>
    public static SessionPreflightResult Collect(IExcelBatch batch, string sessionId)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? worksheet = null;
            dynamic? probeRange = null;
            dynamic? model = null;

            try
            {
                worksheet = ctx.Book.Worksheets.Item[1];
                probeRange = worksheet.Cells[1, 1];
                bool formula2Supported = FormulaCompatibility.SupportsFormula2(batch, probeRange);
                bool readOnly = Convert.ToBoolean(ctx.Book.ReadOnly, CultureInfo.InvariantCulture);
                bool structureProtected = Convert.ToBoolean(ctx.Book.ProtectStructure, CultureInfo.InvariantCulture);
                bool windowsProtected = Convert.ToBoolean(ctx.Book.ProtectWindows, CultureInfo.InvariantCulture);
                bool irmProtected = FileAccessValidator.IsIrmProtected(batch.WorkbookPath);
                string powerPivotStatus;

                try
                {
                    model = ctx.Book.Model;
                    powerPivotStatus = model == null ? "unavailable" : "supported";
                }
                catch (COMException)
                {
                    powerPivotStatus = "unavailable";
                }
                catch (RuntimeBinderException)
                {
                    powerPivotStatus = "unsupported";
                }

                var constraints = new List<string>();
                if (readOnly) constraints.Add("Workbook is read-only.");
                if (structureProtected) constraints.Add("Workbook structure is protected.");
                if (windowsProtected) constraints.Add("Workbook windows are protected.");
                if (irmProtected) constraints.Add("Workbook is protected by IRM or AIP.");

                return new SessionPreflightResult
                {
                    Success = true,
                    SessionId = sessionId,
                    FilePath = batch.WorkbookPath,
                    Excel = new ExcelEnvironmentResult
                    {
                        Version = Convert.ToString(ctx.App.Version, CultureInfo.InvariantCulture) ?? string.Empty,
                        Build = ReadBuild(ctx.App),
                        Bitness = ProcessArchitectureDetector.GetBitness(batch.ExcelProcessId),
                        OperatingSystem = Convert.ToString(ctx.App.OperatingSystem, CultureInfo.InvariantCulture) ?? Environment.OSVersion.VersionString,
                        UiLocale = ReadUiLocale(ctx.App)
                    },
                    Capabilities = new SessionCapabilitiesResult
                    {
                        Formula2 = new Formula2CapabilityResult
                        {
                            Status = formula2Supported ? "supported" : "unsupported",
                            DynamicArrays = formula2Supported
                        },
                        PythonInExcel = new CapabilityResult { Status = "notDetermined" },
                        VbaTrust = new CapabilityResult { Status = VbaCommands.IsVbaTrustEnabled() ? "supported" : "blocked" },
                        PowerPivot = new CapabilityResult { Status = powerPivotStatus }
                    },
                    Workbook = new WorkbookProtectionResult
                    {
                        ReadOnly = readOnly,
                        StructureProtected = structureProtected,
                        WindowsProtected = windowsProtected,
                        IrmProtected = irmProtected
                    },
                    Constraints = constraints,
                    CollectedAtUtc = DateTime.UtcNow
                };
            }
            finally
            {
                ComUtilities.Release(ref model);
                ComUtilities.Release(ref probeRange);
                ComUtilities.Release(ref worksheet);
            }
        });
    }

    private static int ReadBuild(dynamic application) =>
        int.TryParse(Convert.ToString(application.Build, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture, out int build)
            ? build
            : 0;

    private static string ReadUiLocale(dynamic application)
    {
        try
        {
            int lcid = Convert.ToInt32(application.LanguageSettings.LanguageID[2], CultureInfo.InvariantCulture);
            return CultureInfo.GetCultureInfo(lcid).Name;
        }
        catch (COMException)
        {
            return CultureInfo.CurrentUICulture.Name;
        }
        catch (RuntimeBinderException)
        {
            return CultureInfo.CurrentUICulture.Name;
        }
        catch (CultureNotFoundException)
        {
            return CultureInfo.CurrentUICulture.Name;
        }
    }
}
