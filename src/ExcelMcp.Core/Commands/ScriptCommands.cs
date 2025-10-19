using Spectre.Console;
using System.Runtime.InteropServices;
using static ExcelMcp.Core.ExcelHelper;

namespace ExcelMcp.Core.Commands;

/// <summary>
/// VBA script management commands
/// </summary>
public class ScriptCommands : IScriptCommands
{
    /// <summary>
    /// Check if VBA project access is trusted and available
    /// </summary>
    private static bool IsVbaAccessTrusted(string filePath)
    {
        try
        {
            int result = WithExcel(filePath, false, (excel, workbook) =>
            {
                try
                {
                    dynamic vbProject = workbook.VBProject;
                    int componentCount = vbProject.VBComponents.Count; // Try to access VBComponents
                    return 1; // Return 1 for success
                }
                catch (COMException comEx)
                {
                    // Common VBA trust errors
                    if (comEx.ErrorCode == unchecked((int)0x800A03EC)) // Programmatic access not trusted
                    {
                        AnsiConsole.MarkupLine("[red]VBA Error:[/] Programmatic access to VBA project is not trusted");
                        AnsiConsole.MarkupLine("[yellow]Solution:[/] Run: [cyan]ExcelCLI setup-vba-trust[/]");
                    }
                    else
                    {
                        AnsiConsole.MarkupLine($"[red]VBA COM Error:[/] 0x{comEx.ErrorCode:X8} - {comEx.Message.EscapeMarkup()}");
                    }
                    return 0;
                }
                catch (Exception ex)
                {
                    AnsiConsole.MarkupLine($"[red]VBA Access Error:[/] {ex.Message.EscapeMarkup()}");
                    return 0;
                }
            });
            return result == 1;
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error checking VBA access:[/] {ex.Message.EscapeMarkup()}");
            return false;
        }
    }

    /// <summary>
    /// Validate that file is macro-enabled (.xlsm) for VBA operations
    /// </summary>
    private static bool ValidateVbaFile(string filePath)
    {
        string extension = Path.GetExtension(filePath).ToLowerInvariant();
        if (extension != ".xlsm")
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] VBA operations require macro-enabled workbooks (.xlsm)");
            AnsiConsole.MarkupLine($"[yellow]Current file:[/] {Path.GetFileName(filePath)} ({extension})");
            AnsiConsole.MarkupLine($"[yellow]Solutions:[/]");
            AnsiConsole.MarkupLine($"  • Create new .xlsm file: [cyan]ExcelCLI create-empty \"file.xlsm\"[/]");
            AnsiConsole.MarkupLine($"  • Save existing file as .xlsm in Excel");
            AnsiConsole.MarkupLine($"  • Convert with: [cyan]ExcelCLI sheet-copy \"{filePath}\" \"Sheet1\" \"newfile.xlsm\"[/]");
            return false;
        }
        return true;
    }

    public int List(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] script-list <file.xlsx>");
            return 1;
        }

        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[bold]Office Scripts in:[/] {Path.GetFileName(args[1])}\n");

        return WithExcel(args[1], false, (excel, workbook) =>
        {
            try
            {
                var scripts = new List<(string Name, string Type)>();

                // Try to access VBA project
                try
                {
                    dynamic vbaProject = workbook.VBProject;
                    dynamic vbComponents = vbaProject.VBComponents;

                    for (int i = 1; i <= vbComponents.Count; i++)
                    {
                        dynamic component = vbComponents.Item(i);
                        string name = component.Name;
                        int type = component.Type;

                        string typeStr = type switch
                        {
                            1 => "Module",
                            2 => "Class",
                            3 => "Form",
                            100 => "Document",
                            _ => $"Type{type}"
                        };

                        scripts.Add((name, typeStr));
                    }
                }
                catch
                {
                    AnsiConsole.MarkupLine("[yellow]Note:[/] VBA macros not accessible or not present");
                }

                // Display scripts
                if (scripts.Count > 0)
                {
                    var table = new Table();
                    table.AddColumn("[bold]Script Name[/]");
                    table.AddColumn("[bold]Type[/]");

                    foreach (var (name, type) in scripts.OrderBy(s => s.Name))
                    {
                        table.AddRow(name.EscapeMarkup(), type.EscapeMarkup());
                    }

                    AnsiConsole.Write(table);
                    AnsiConsole.MarkupLine($"\n[dim]Total: {scripts.Count} script(s)[/]");
                }
                else
                {
                    AnsiConsole.MarkupLine("[yellow]No VBA scripts found[/]");
                    AnsiConsole.MarkupLine("[dim]Note: Office Scripts (.ts) are not stored in Excel files[/]");
                }

                return 0;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
                return 1;
            }
        });
    }

    public int Export(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] script-export <file.xlsx> <script-name> <output-file>");
            return 1;
        }

        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        string scriptName = args[2];
        string outputFile = args.Length > 3 ? args[3] : $"{scriptName}.vba";

        return WithExcel(args[1], false, (excel, workbook) =>
        {
            try
            {
                dynamic vbaProject = workbook.VBProject;
                dynamic vbComponents = vbaProject.VBComponents;
                dynamic? targetComponent = null;

                for (int i = 1; i <= vbComponents.Count; i++)
                {
                    dynamic component = vbComponents.Item(i);
                    if (component.Name == scriptName)
                    {
                        targetComponent = component;
                        break;
                    }
                }

                if (targetComponent == null)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] Script '{scriptName}' not found");
                    return 1;
                }

                // Get the code module
                dynamic codeModule = targetComponent.CodeModule;
                int lineCount = codeModule.CountOfLines;

                if (lineCount > 0)
                {
                    string code = codeModule.Lines(1, lineCount);
                    File.WriteAllText(outputFile, code);

                    AnsiConsole.MarkupLine($"[green]√[/] Exported script '{scriptName}' to '{outputFile}'");
                    AnsiConsole.MarkupLine($"[dim]{lineCount} lines[/]");
                    return 0;
                }
                else
                {
                    AnsiConsole.MarkupLine($"[yellow]Warning:[/] Script '{scriptName}' is empty");
                    return 1;
                }
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
                AnsiConsole.MarkupLine("[yellow]Tip:[/] Make sure 'Trust access to the VBA project object model' is enabled");
                return 1;
            }
        });
    }

    public int Run(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] script-run <file.xlsm> <macro-name> [[param1]] [[param2]] ...");
            AnsiConsole.MarkupLine("[yellow]Example:[/] script-run \"Plan.xlsm\" \"ProcessData\"");
            AnsiConsole.MarkupLine("[yellow]Example:[/] script-run \"Plan.xlsm\" \"CalculateTotal\" \"Sheet1\" \"A1:C10\"");
            return 1;
        }

        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        
        // Validate file format
        if (!ValidateVbaFile(filePath))
        {
            return 1;
        }

        string macroName = args[2];
        var parameters = args.Skip(3).ToArray();

        return WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                AnsiConsole.MarkupLine($"[cyan]Running macro:[/] {macroName}");
                if (parameters.Length > 0)
                {
                    AnsiConsole.MarkupLine($"[dim]Parameters: {string.Join(", ", parameters)}[/]");
                }

                // Prepare parameters for Application.Run
                object[] runParams = new object[31]; // Application.Run supports up to 30 parameters + macro name
                runParams[0] = macroName;
                
                for (int i = 0; i < Math.Min(parameters.Length, 30); i++)
                {
                    runParams[i + 1] = parameters[i];
                }
                
                // Fill remaining parameters with missing values
                for (int i = parameters.Length + 1; i < 31; i++)
                {
                    runParams[i] = Type.Missing;
                }

                // Execute the macro
                dynamic result = excel.Run(
                    runParams[0], runParams[1], runParams[2], runParams[3], runParams[4],
                    runParams[5], runParams[6], runParams[7], runParams[8], runParams[9],
                    runParams[10], runParams[11], runParams[12], runParams[13], runParams[14],
                    runParams[15], runParams[16], runParams[17], runParams[18], runParams[19],
                    runParams[20], runParams[21], runParams[22], runParams[23], runParams[24],
                    runParams[25], runParams[26], runParams[27], runParams[28], runParams[29],
                    runParams[30]
                );

                AnsiConsole.MarkupLine($"[green]√[/] Macro '{macroName}' completed successfully");
                
                // Display result if macro returned something
                if (result != null && result != Type.Missing)
                {
                    AnsiConsole.MarkupLine($"[cyan]Result:[/] {result.ToString().EscapeMarkup()}");
                }

                return 0;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
                
                if (ex.Message.Contains("macro") || ex.Message.Contains("procedure"))
                {
                    AnsiConsole.MarkupLine("[yellow]Tip:[/] Make sure the macro name is correct and the VBA code is present");
                    AnsiConsole.MarkupLine("[yellow]Tip:[/] Use 'script-list' to see available VBA modules and procedures");
                }
                
                return 1;
            }
        });
    }

    /// <summary>
    /// Import VBA code from file into Excel workbook
    /// </summary>
    public async Task<int> Import(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] script-import <file.xlsm> <module-name> <vba-file>");
            AnsiConsole.MarkupLine("[yellow]Note:[/] VBA operations require macro-enabled workbooks (.xlsm)");
            return 1;
        }

        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        if (!File.Exists(args[3]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] VBA file not found: {args[3]}");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        
        // Validate file format
        if (!ValidateVbaFile(filePath))
        {
            return 1;
        }
        
        // Check VBA access first
        if (!IsVbaAccessTrusted(filePath))
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Programmatic access to Visual Basic Project is not trusted");
            AnsiConsole.MarkupLine("[yellow]Tip:[/] Make sure 'Trust access to the VBA project object model' is enabled in Excel");
            AnsiConsole.MarkupLine("[yellow]Tip:[/] File → Options → Trust Center → Trust Center Settings → Macro Settings");
            return 1;
        }

        string moduleName = args[2];
        string vbaFilePath = args[3];

        try
        {
            string vbaCode = await File.ReadAllTextAsync(vbaFilePath);
            
            return WithExcel(filePath, true, (excel, workbook) =>
            {
                try
                {
                    // Access the VBA project
                    dynamic vbProject = workbook.VBProject;
                    dynamic vbComponents = vbProject.VBComponents;

                    // Check if module already exists
                    dynamic? existingModule = null;
                    try
                    {
                        existingModule = vbComponents.Item(moduleName);
                    }
                    catch
                    {
                        // Module doesn't exist, which is fine for import
                    }

                    if (existingModule != null)
                    {
                        AnsiConsole.MarkupLine($"[yellow]Warning:[/] Module '{moduleName}' already exists. Use 'script-update' to modify existing modules.");
                        return 1;
                    }

                    // Add new module
                    const int vbext_ct_StdModule = 1;
                    dynamic newModule = vbComponents.Add(vbext_ct_StdModule);
                    newModule.Name = moduleName;

                    // Add the VBA code
                    dynamic codeModule = newModule.CodeModule;
                    codeModule.AddFromString(vbaCode);

                    // Force save to ensure the module is persisted
                    workbook.Save();

                    AnsiConsole.MarkupLine($"[green]✓[/] Imported VBA module '{moduleName}'");
                    return 0;
                }
                catch (Exception ex)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
                    
                    if (ex.Message.Contains("access") || ex.Message.Contains("trust"))
                    {
                        AnsiConsole.MarkupLine("[yellow]Tip:[/] Make sure 'Trust access to the VBA project object model' is enabled in Excel");
                        AnsiConsole.MarkupLine("[yellow]Tip:[/] File → Options → Trust Center → Trust Center Settings → Macro Settings");
                    }
                    
                    return 1;
                }
            });
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error reading VBA file:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }
    }

    /// <summary>
    /// Update existing VBA module with new code from file
    /// </summary>
    public async Task<int> Update(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] script-update <file.xlsm> <module-name> <vba-file>");
            AnsiConsole.MarkupLine("[yellow]Note:[/] VBA operations require macro-enabled workbooks (.xlsm)");
            return 1;
        }

        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        if (!File.Exists(args[3]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] VBA file not found: {args[3]}");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        
        // Validate file format
        if (!ValidateVbaFile(filePath))
        {
            return 1;
        }
        
        // Check VBA access first
        if (!IsVbaAccessTrusted(filePath))
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Programmatic access to Visual Basic Project is not trusted");
            AnsiConsole.MarkupLine("[yellow]Tip:[/] Make sure 'Trust access to the VBA project object model' is enabled in Excel");
            AnsiConsole.MarkupLine("[yellow]Tip:[/] File → Options → Trust Center → Trust Center Settings → Macro Settings");
            return 1;
        }
        
        string moduleName = args[2];
        string vbaFilePath = args[3];

        try
        {
            string vbaCode = await File.ReadAllTextAsync(vbaFilePath);
            
            return WithExcel(filePath, true, (excel, workbook) =>
            {
                try
                {
                    // Access the VBA project
                    dynamic vbProject = workbook.VBProject;
                    dynamic vbComponents = vbProject.VBComponents;

                    // Find the existing module
                    dynamic? targetModule = null;
                    try
                    {
                        targetModule = vbComponents.Item(moduleName);
                    }
                    catch
                    {
                        AnsiConsole.MarkupLine($"[red]Error:[/] Module '{moduleName}' not found. Use 'script-import' to create new modules.");
                        return 1;
                    }

                    // Clear existing code and add new code
                    dynamic codeModule = targetModule.CodeModule;
                    int lineCount = codeModule.CountOfLines;
                    if (lineCount > 0)
                    {
                        codeModule.DeleteLines(1, lineCount);
                    }
                    codeModule.AddFromString(vbaCode);

                    // Force save to ensure the changes are persisted
                    workbook.Save();

                    AnsiConsole.MarkupLine($"[green]✓[/] Updated VBA module '{moduleName}'");
                    return 0;
                }
                catch (Exception ex)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
                    
                    if (ex.Message.Contains("access") || ex.Message.Contains("trust"))
                    {
                        AnsiConsole.MarkupLine("[yellow]Tip:[/] Make sure 'Trust access to the VBA project object model' is enabled in Excel");
                        AnsiConsole.MarkupLine("[yellow]Tip:[/] File → Options → Trust Center → Trust Center Settings → Macro Settings");
                    }
                    
                    return 1;
                }
            });
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error reading VBA file:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }
    }
}
