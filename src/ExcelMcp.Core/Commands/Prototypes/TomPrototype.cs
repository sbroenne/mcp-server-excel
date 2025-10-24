using Microsoft.AnalysisServices.Tabular;

namespace Sbroenne.ExcelMcp.Core.Commands.Prototypes;

/// <summary>
/// Prototype for testing TOM API with Excel Data Models.
/// This class validates that Microsoft.AnalysisServices.NetCore.retail.amd64 works with Excel files.
/// </summary>
public class TomPrototype
{
    /// <summary>
    /// Test connection to Excel Data Model using TOM API.
    /// </summary>
    public static bool CanConnectToExcelDataModel(string excelFilePath)
    {
        Server? server = null;
        try
        {
            server = new Server();
            
            // Try different connection string formats for Excel
            string[] connectionFormats = new[]
            {
                $"Provider=MSOLAP;Data Source={excelFilePath};",
                $"Data Source={excelFilePath};",
                $"Provider=MSOLAP.8;Data Source={excelFilePath};",
                $"DataSource={excelFilePath};Provider=MSOLAP;"
            };

            foreach (var connString in connectionFormats)
            {
                try
                {
                    server.Connect(connString);
                    
                    if (server.Connected)
                    {
                        Console.WriteLine($"✅ Connected with: {connString}");
                        Console.WriteLine($"   Server Version: {server.Version}");
                        Console.WriteLine($"   Databases: {server.Databases.Count}");
                        
                        if (server.Databases.Count > 0)
                        {
                            Database db = server.Databases[0];
                            Console.WriteLine($"   Database Name: {db.Name}");
                            Console.WriteLine($"   Database ID: {db.ID}");
                            
                            if (db.Model != null)
                            {
                                Console.WriteLine($"   Model Tables: {db.Model.Tables.Count}");
                                
                                // Count measures across all tables
                                int totalMeasures = 0;
                                foreach (Table table in db.Model.Tables)
                                {
                                    totalMeasures += table.Measures.Count;
                                }
                                Console.WriteLine($"   Model Measures: {totalMeasures}");
                                Console.WriteLine($"   Model Relationships: {db.Model.Relationships.Count}");
                            }
                        }
                        
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Failed with: {connString}");
                    Console.WriteLine($"   Error: {ex.Message}");
                }
                finally
                {
                    if (server.Connected)
                    {
                        server.Disconnect();
                    }
                }
            }
            
            return false;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ TOM API Error: {ex.Message}");
            Console.WriteLine($"   Type: {ex.GetType().Name}");
            if (ex.InnerException != null)
            {
                Console.WriteLine($"   Inner: {ex.InnerException.Message}");
            }
            return false;
        }
        finally
        {
            if (server?.Connected == true)
            {
                server.Disconnect();
            }
        }
    }

    /// <summary>
    /// Test creating a measure using TOM API.
    /// </summary>
    public static bool CanCreateMeasure(string excelFilePath, string tableName, string measureName, string daxFormula)
    {
        Server? server = null;
        try
        {
            server = new Server();
            server.Connect($"Provider=MSOLAP;Data Source={excelFilePath};");

            if (!server.Connected || server.Databases.Count == 0)
            {
                Console.WriteLine("❌ Not connected or no database found");
                return false;
            }

            Database db = server.Databases[0];
            Model model = db.Model;

            // Find table
            Table? table = model.Tables.Find(tableName);
            if (table == null)
            {
                Console.WriteLine($"❌ Table '{tableName}' not found");
                return false;
            }

            // Create measure
            Measure newMeasure = new Measure
            {
                Name = measureName,
                Expression = daxFormula,
                Description = "Created via TOM API prototype"
            };

            table.Measures.Add(newMeasure);

            // Save changes
            model.SaveChanges();

            Console.WriteLine($"✅ Measure '{measureName}' created successfully");
            Console.WriteLine($"   Table: {tableName}");
            Console.WriteLine($"   Formula: {daxFormula}");

            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Create Measure Error: {ex.Message}");
            Console.WriteLine($"   Type: {ex.GetType().Name}");
            if (ex.InnerException != null)
            {
                Console.WriteLine($"   Inner: {ex.InnerException.Message}");
            }
            return false;
        }
        finally
        {
            if (server?.Connected == true)
            {
                server.Disconnect();
            }
        }
    }

    /// <summary>
    /// Test creating a relationship using TOM API.
    /// </summary>
    public static bool CanCreateRelationship(
        string excelFilePath,
        string fromTableName,
        string fromColumnName,
        string toTableName,
        string toColumnName)
    {
        Server? server = null;
        try
        {
            server = new Server();
            server.Connect($"Provider=MSOLAP;Data Source={excelFilePath};");

            if (!server.Connected || server.Databases.Count == 0)
            {
                Console.WriteLine("❌ Not connected or no database found");
                return false;
            }

            Database db = server.Databases[0];
            Model model = db.Model;

            // Find tables and columns
            Table? fromTable = model.Tables.Find(fromTableName);
            Table? toTable = model.Tables.Find(toTableName);

            if (fromTable == null || toTable == null)
            {
                Console.WriteLine($"❌ Table not found: {fromTableName} or {toTableName}");
                return false;
            }

            Column? fromColumn = fromTable.Columns.Find(fromColumnName);
            Column? toColumn = toTable.Columns.Find(toColumnName);

            if (fromColumn == null || toColumn == null)
            {
                Console.WriteLine($"❌ Column not found: {fromColumnName} or {toColumnName}");
                return false;
            }

            // Create relationship
            SingleColumnRelationship relationship = new SingleColumnRelationship
            {
                Name = $"{fromTableName}_{fromColumnName}_to_{toTableName}_{toColumnName}",
                FromColumn = fromColumn,
                ToColumn = toColumn,
                FromCardinality = RelationshipEndCardinality.Many,
                ToCardinality = RelationshipEndCardinality.One,
                IsActive = true
            };

            model.Relationships.Add(relationship);

            // Save changes
            model.SaveChanges();

            Console.WriteLine($"✅ Relationship created successfully");
            Console.WriteLine($"   From: {fromTableName}.{fromColumnName}");
            Console.WriteLine($"   To: {toTableName}.{toColumnName}");

            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Create Relationship Error: {ex.Message}");
            Console.WriteLine($"   Type: {ex.GetType().Name}");
            if (ex.InnerException != null)
            {
                Console.WriteLine($"   Inner: {ex.InnerException.Message}");
            }
            return false;
        }
        finally
        {
            if (server?.Connected == true)
            {
                server.Disconnect();
            }
        }
    }

    /// <summary>
    /// List all available TOM namespaces and types for research.
    /// </summary>
    public static void ListTomCapabilities()
    {
        Console.WriteLine("📦 TOM API Capabilities:");
        Console.WriteLine();
        
        Console.WriteLine("Core Types:");
        Console.WriteLine("  - Server: Connection to Analysis Services");
        Console.WriteLine("  - Database: Excel workbook Data Model");
        Console.WriteLine("  - Model: Contains tables, measures, relationships");
        Console.WriteLine("  - Table: Data Model table");
        Console.WriteLine("  - Column: Table column");
        Console.WriteLine("  - Measure: DAX measure");
        Console.WriteLine("  - Relationship: Table relationship");
        Console.WriteLine();
        
        Console.WriteLine("Available Operations:");
        Console.WriteLine("  ✅ Connect to Excel Data Model");
        Console.WriteLine("  ✅ Read Model metadata (tables, measures, relationships)");
        Console.WriteLine("  ✅ Create new measures");
        Console.WriteLine("  ✅ Update existing measures");
        Console.WriteLine("  ✅ Delete measures");
        Console.WriteLine("  ✅ Create relationships");
        Console.WriteLine("  ✅ Update relationships");
        Console.WriteLine("  ✅ Delete relationships");
        Console.WriteLine("  ✅ SaveChanges() to persist to Excel file");
        Console.WriteLine();
        
        Console.WriteLine("Package Info:");
        Console.WriteLine("  - Package: Microsoft.AnalysisServices.NetCore.retail.amd64");
        Console.WriteLine("  - Version: 19.84.1");
        Console.WriteLine("  - Target Framework: .NET Core / .NET 5+");
        Console.WriteLine("  - Compatible with: .NET 9.0 ✅");
    }
}
