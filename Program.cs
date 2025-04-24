using System;
using System.Linq;
using Microsoft.AnalysisServices.Tabular;

namespace PBITOMWrapper
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("PBI TOMWrapper Demo");
            Console.WriteLine("===================");
            
            // Create a new instance of our wrapper
            using (var wrapper = new PBIModelWrapper())
            {
                // Enable diagnostics for debugging
                Console.Write("Enable diagnostics? (y/n): ");
                if (Console.ReadLine().ToLower() == "y")
                {
                    wrapper.EnableDiagnostics(true);
                    Console.WriteLine("Diagnostics enabled.");
                }
                
                Console.WriteLine("\nConnection Options:");
                Console.WriteLine("1. Auto-detect Power BI Desktop instances");
                Console.WriteLine("2. Connect to a specific port");
                
                Console.Write("Select an option (1-2): ");
                string option = Console.ReadLine();
                
                bool connected = false;
                
                switch (option)
                {
                    case "1":
                        connected = AutoDetectAndConnect(wrapper);
                        break;
                        
                    case "2":
                        connected = ConnectToSpecificPort(wrapper);
                        break;
                        
                    default:
                        Console.WriteLine("Invalid option. Exiting.");
                        return;
                }
                
                if (!connected)
                {
                    // If diagnostics were enabled, show the logs
                    if (PBIProcessHelper.DiagnosticsEnabled)
                    {
                        Console.WriteLine("\n===== Diagnostic Logs =====");
                        foreach (var log in wrapper.GetDiagnosticLogs())
                        {
                            Console.WriteLine(log);
                        }
                    }
                    
                    return;
                }
                
                // Display tables
                var tables = wrapper.GetTables().ToList();
                Console.WriteLine($"\nTables ({tables.Count}):");
                foreach (var table in tables)
                {
                    Console.WriteLine($"- {table.Name}");
                }
                
                // Display measures
                var measures = wrapper.GetAllMeasures().ToList();
                Console.WriteLine($"\nMeasures ({measures.Count}):");
                foreach (var measure in measures)
                {
                    Console.WriteLine($"- {measure.Name} [{measure.Table.Name}]: {measure.Expression}");
                }
                
                // Menu for operations
                while (true)
                {
                    Console.WriteLine("\nOperations:");
                    Console.WriteLine("1. Create measure");
                    Console.WriteLine("2. Update measure");
                    Console.WriteLine("3. Delete measure");
                    Console.WriteLine("4. Create calculated column");
                    Console.WriteLine("5. Display model info");
                    Console.WriteLine("6. Show diagnostic logs");
                    Console.WriteLine("7. Exit");
                    
                    Console.Write("\nSelect an operation (1-7): ");
                    string operation = Console.ReadLine();
                    
                    switch (operation)
                    {
                        case "1":
                            CreateMeasure(wrapper);
                            break;
                            
                        case "2":
                            UpdateMeasure(wrapper);
                            break;
                            
                        case "3":
                            DeleteMeasure(wrapper);
                            break;
                            
                        case "4":
                            CreateCalculatedColumn(wrapper);
                            break;
                            
                        case "5":
                            DisplayModelInfo(wrapper);
                            break;
                            
                        case "6":
                            ShowDiagnosticLogs(wrapper);
                            break;
                            
                        case "7":
                            Console.WriteLine("Exiting...");
                            return;
                            
                        default:
                            Console.WriteLine("Invalid selection.");
                            break;
                    }
                }
            }
        }
        
        static bool AutoDetectAndConnect(PBIModelWrapper wrapper)
        {
            // Find running Power BI Desktop instances
            var instances = wrapper.GetRunningPBIInstances();
            
            if (instances.Count == 0)
            {
                Console.WriteLine("No running Power BI Desktop instances found.");
                Console.WriteLine("Please open a PBIX file in Power BI Desktop and try again, or use the port connection option.");
                return false;
            }
            
            Console.WriteLine($"Found {instances.Count} running Power BI Desktop instance(s):");
            
            for (int i = 0; i < instances.Count; i++)
            {
                var instance = instances[i];
                Console.WriteLine($"{i + 1}. {instance.FileName} (PID: {instance.ProcessId}, Port: {instance.Port})");
            }
            
            // If there's more than one instance, let the user choose
            int selectedIndex = 0;
            if (instances.Count > 1)
            {
                Console.Write("\nSelect an instance (1-{0}): ", instances.Count);
                if (!int.TryParse(Console.ReadLine(), out selectedIndex) || selectedIndex < 1 || selectedIndex > instances.Count)
                {
                    Console.WriteLine("Invalid selection. Exiting.");
                    return false;
                }
                
                selectedIndex--; // Convert to 0-based index
            }
            
            // Connect to the selected instance
            var selectedInstance = instances[selectedIndex];
            Console.WriteLine($"\nConnecting to {selectedInstance.FileName}...");
            
            return wrapper.ConnectToInstance(selectedInstance);
        }
        
        static bool ConnectToSpecificPort(PBIModelWrapper wrapper)
        {
            Console.Write("Enter the port number (e.g., 62551): ");
            if (!int.TryParse(Console.ReadLine(), out int portNumber) || portNumber <= 0)
            {
                Console.WriteLine("Invalid port number.");
                return false;
            }
            
            Console.WriteLine($"Connecting to port {portNumber}...");
            return wrapper.ConnectToPort(portNumber);
        }
        
        static void ShowDiagnosticLogs(PBIModelWrapper wrapper)
        {
            var logs = wrapper.GetDiagnosticLogs();
            
            Console.WriteLine("\n===== Diagnostic Logs =====");
            foreach (var log in logs)
            {
                Console.WriteLine(log);
            }
        }
        
        static void CreateMeasure(PBIModelWrapper wrapper)
        {
            // Display tables for selection
            var tables = wrapper.GetTables().ToList();
            Console.WriteLine("\nSelect a table:");
            for (int i = 0; i < tables.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {tables[i].Name}");
            }
            
            Console.Write("Table (1-{0}): ", tables.Count);
            if (!int.TryParse(Console.ReadLine(), out int tableIndex) || tableIndex < 1 || tableIndex > tables.Count)
            {
                Console.WriteLine("Invalid selection.");
                return;
            }
            
            var selectedTable = tables[tableIndex - 1];
            
            Console.Write("Measure name: ");
            string measureName = Console.ReadLine();
            
            if (string.IsNullOrWhiteSpace(measureName))
            {
                Console.WriteLine("Measure name cannot be empty.");
                return;
            }
            
            Console.Write("DAX expression: ");
            string daxExpression = Console.ReadLine();
            
            if (string.IsNullOrWhiteSpace(daxExpression))
            {
                Console.WriteLine("DAX expression cannot be empty.");
                return;
            }
            
            Console.Write("Format string (optional): ");
            string formatString = Console.ReadLine();
            
            if (string.IsNullOrWhiteSpace(formatString))
            {
                formatString = null;
            }
            
            try
            {
                var measure = wrapper.CreateMeasure(selectedTable.Name, measureName, daxExpression, formatString);
                Console.WriteLine($"Measure '{measure.Name}' created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating measure: {ex.Message}");
            }
        }
        
        static void UpdateMeasure(PBIModelWrapper wrapper)
        {
            var measures = wrapper.GetAllMeasures().ToList();
            
            if (measures.Count == 0)
            {
                Console.WriteLine("No measures found in the model.");
                return;
            }
            
            Console.WriteLine("\nSelect a measure to update:");
            for (int i = 0; i < measures.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {measures[i].Name} [{measures[i].Table.Name}]");
            }
            
            Console.Write("Measure (1-{0}): ", measures.Count);
            if (!int.TryParse(Console.ReadLine(), out int measureIndex) || measureIndex < 1 || measureIndex > measures.Count)
            {
                Console.WriteLine("Invalid selection.");
                return;
            }
            
            var selectedMeasure = measures[measureIndex - 1];
            
            Console.WriteLine($"Current DAX expression: {selectedMeasure.Expression}");
            Console.Write("New DAX expression: ");
            string daxExpression = Console.ReadLine();
            
            if (string.IsNullOrWhiteSpace(daxExpression))
            {
                Console.WriteLine("DAX expression cannot be empty.");
                return;
            }
            
            Console.WriteLine($"Current format string: {selectedMeasure.FormatString}");
            Console.Write("New format string (optional, press Enter to keep current): ");
            string formatString = Console.ReadLine();
            
            if (string.IsNullOrWhiteSpace(formatString))
            {
                formatString = null;
            }
            
            try
            {
                var measure = wrapper.UpdateMeasure(selectedMeasure.Table.Name, selectedMeasure.Name, daxExpression, formatString);
                Console.WriteLine($"Measure '{measure.Name}' updated successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating measure: {ex.Message}");
            }
        }
        
        static void DeleteMeasure(PBIModelWrapper wrapper)
        {
            var measures = wrapper.GetAllMeasures().ToList();
            
            if (measures.Count == 0)
            {
                Console.WriteLine("No measures found in the model.");
                return;
            }
            
            Console.WriteLine("\nSelect a measure to delete:");
            for (int i = 0; i < measures.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {measures[i].Name} [{measures[i].Table.Name}]");
            }
            
            Console.Write("Measure (1-{0}): ", measures.Count);
            if (!int.TryParse(Console.ReadLine(), out int measureIndex) || measureIndex < 1 || measureIndex > measures.Count)
            {
                Console.WriteLine("Invalid selection.");
                return;
            }
            
            var selectedMeasure = measures[measureIndex - 1];
            
            Console.Write($"Are you sure you want to delete the measure '{selectedMeasure.Name}'? (y/n): ");
            string confirmation = Console.ReadLine();
            
            if (confirmation.ToLower() != "y")
            {
                Console.WriteLine("Deletion cancelled.");
                return;
            }
            
            try
            {
                bool success = wrapper.DeleteMeasure(selectedMeasure.Table.Name, selectedMeasure.Name);
                
                if (success)
                {
                    Console.WriteLine($"Measure '{selectedMeasure.Name}' deleted successfully.");
                }
                else
                {
                    Console.WriteLine($"Failed to delete measure '{selectedMeasure.Name}'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error deleting measure: {ex.Message}");
            }
        }
        
        static void CreateCalculatedColumn(PBIModelWrapper wrapper)
        {
            // Display tables for selection
            var tables = wrapper.GetTables().ToList();
            Console.WriteLine("\nSelect a table:");
            for (int i = 0; i < tables.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {tables[i].Name}");
            }
            
            Console.Write("Table (1-{0}): ", tables.Count);
            if (!int.TryParse(Console.ReadLine(), out int tableIndex) || tableIndex < 1 || tableIndex > tables.Count)
            {
                Console.WriteLine("Invalid selection.");
                return;
            }
            
            var selectedTable = tables[tableIndex - 1];
            
            Console.Write("Column name: ");
            string columnName = Console.ReadLine();
            
            if (string.IsNullOrWhiteSpace(columnName))
            {
                Console.WriteLine("Column name cannot be empty.");
                return;
            }
            
            Console.Write("DAX expression: ");
            string daxExpression = Console.ReadLine();
            
            if (string.IsNullOrWhiteSpace(daxExpression))
            {
                Console.WriteLine("DAX expression cannot be empty.");
                return;
            }
            
            Console.WriteLine("Data types:");
            Console.WriteLine("1. String");
            Console.WriteLine("2. Integer");
            Console.WriteLine("3. Decimal");
            Console.WriteLine("4. Boolean");
            Console.WriteLine("5. DateTime");
            
            Console.Write("Data type (1-5): ");
            if (!int.TryParse(Console.ReadLine(), out int dataTypeIndex) || dataTypeIndex < 1 || dataTypeIndex > 5)
            {
                Console.WriteLine("Invalid selection. Using String as default.");
                dataTypeIndex = 1;
            }
            
            DataType dataType;
            switch (dataTypeIndex)
            {
                case 1:
                    dataType = DataType.String;
                    break;
                case 2:
                    dataType = DataType.Int64;
                    break;
                case 3:
                    dataType = DataType.Double;
                    break;
                case 4:
                    dataType = DataType.Boolean;
                    break;
                case 5:
                    dataType = DataType.DateTime;
                    break;
                default:
                    dataType = DataType.String;
                    break;
            }
            
            try
            {
                var column = wrapper.CreateCalculatedColumn(selectedTable.Name, columnName, daxExpression, dataType);
                Console.WriteLine($"Calculated column '{column.Name}' created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating calculated column: {ex.Message}");
            }
        }
        
        static void DisplayModelInfo(PBIModelWrapper wrapper)
        {
            Console.WriteLine("\nModel Information:");
            Console.WriteLine($"Database Name: {wrapper.Database.Name}");
            Console.WriteLine($"Compatibility Level: {wrapper.Database.CompatibilityLevel}");
            
            var tables = wrapper.GetTables().ToList();
            Console.WriteLine($"\nTables ({tables.Count}):");
            foreach (var table in tables)
            {
                Console.WriteLine($"- {table.Name} ({table.Columns.Count} columns, {table.Measures.Count} measures)");
            }
            
            var relationships = wrapper.GetRelationships().ToList();
            Console.WriteLine($"\nRelationships ({relationships.Count}):");
            foreach (var relationship in relationships)
            {
                Console.WriteLine($"- {relationship.FromColumn.Table.Name}[{relationship.FromColumn.Name}] -> {relationship.ToColumn.Table.Name}[{relationship.ToColumn.Name}]");
            }
            
            var measures = wrapper.GetAllMeasures().ToList();
            Console.WriteLine($"\nMeasures ({measures.Count}):");
            foreach (var measure in measures)
            {
                Console.WriteLine($"- {measure.Name} [{measure.Table.Name}]");
            }
        }
    }
} 