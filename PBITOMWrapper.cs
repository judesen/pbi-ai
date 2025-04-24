using System;
using System.Linq;
using System.Collections.Generic;
using Microsoft.AnalysisServices.Tabular;

namespace PBITOMWrapper
{
    /// <summary>
    /// A wrapper class for interacting with local Power BI Desktop models using the Tabular Object Model (TOM).
    /// </summary>
    public class PBIModelWrapper : IDisposable
    {
        private Server _server;
        private Database _database;
        private Model _model;
        private bool _isConnected = false;
        private PBIProcessHelper.PBIInstance _currentInstance;

        /// <summary>
        /// Gets the underlying TOM Server object.
        /// </summary>
        public Server Server => _server;

        /// <summary>
        /// Gets the underlying TOM Database object.
        /// </summary>
        public Database Database => _database;

        /// <summary>
        /// Gets the underlying TOM Model object.
        /// </summary>
        public Model Model => _model;

        /// <summary>
        /// Indicates whether the wrapper is connected to a Power BI Desktop model.
        /// </summary>
        public bool IsConnected => _isConnected;

        /// <summary>
        /// Gets information about the current Power BI Desktop instance.
        /// </summary>
        public PBIProcessHelper.PBIInstance CurrentInstance => _currentInstance;

        /// <summary>
        /// Creates a new instance of the PBIModelWrapper class.
        /// </summary>
        public PBIModelWrapper()
        {
            _server = new Server();
        }

        /// <summary>
        /// Enables or disables diagnostics logging.
        /// </summary>
        /// <param name="enabled">Whether diagnostics logging should be enabled.</param>
        public void EnableDiagnostics(bool enabled)
        {
            PBIProcessHelper.DiagnosticsEnabled = enabled;
        }

        /// <summary>
        /// Gets a list of running Power BI Desktop instances.
        /// </summary>
        /// <returns>A list of running Power BI Desktop instances.</returns>
        public List<PBIProcessHelper.PBIInstance> GetRunningPBIInstances()
        {
            return PBIProcessHelper.GetRunningPBIInstances();
        }

        /// <summary>
        /// Connects to a specific Power BI Desktop instance.
        /// </summary>
        /// <param name="instance">The Power BI Desktop instance to connect to.</param>
        /// <returns>True if connection was successful, false otherwise.</returns>
        public bool ConnectToInstance(PBIProcessHelper.PBIInstance instance)
        {
            if (instance == null)
                throw new ArgumentNullException(nameof(instance));

            if (instance.Port <= 0)
                throw new ArgumentException("Invalid port number for Power BI instance.");

            try
            {
                // Create the connection string
                string connectionString = $"Data Source=localhost:{instance.Port}";
                
                // Disconnect if already connected
                if (_isConnected)
                {
                    _server.Disconnect();
                }
                
                // Connect to the Analysis Services instance
                _server.Connect(connectionString);

                // If connection successful, get the first database (there should only be one)
                if (_server.Databases.Count > 0)
                {
                    _database = _server.Databases[0];
                    _model = _database.Model;
                    _isConnected = true;
                    _currentInstance = instance;
                    
                    Console.WriteLine($"Successfully connected to Power BI Desktop instance on port {instance.Port}.");
                    Console.WriteLine($"Database: {_database.Name}, Compatibility Level: {_database.CompatibilityLevel}");
                    
                    return true;
                }
                else
                {
                    throw new Exception("No databases found in the Power BI Desktop instance.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error connecting to Power BI Desktop: {ex.Message}");
                _isConnected = false;
                _currentInstance = null;
                return false;
            }
        }

        /// <summary>
        /// Connects to a Power BI Desktop instance on a specific port.
        /// </summary>
        /// <param name="port">The port to connect to.</param>
        /// <returns>True if connection was successful, false otherwise.</returns>
        public bool ConnectToPort(int port)
        {
            if (port <= 0)
                throw new ArgumentException("Port number must be greater than zero.", nameof(port));

            try
            {
                var instance = PBIProcessHelper.CreateInstanceWithPort(port);
                return ConnectToInstance(instance);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error connecting to port {port}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Connects to the first available Power BI Desktop instance.
        /// </summary>
        /// <returns>True if connection was successful, false otherwise.</returns>
        public bool ConnectToLocalPBIDesktop()
        {
            var instances = GetRunningPBIInstances();
            if (instances.Count == 0)
            {
                Console.WriteLine("No running Power BI Desktop instances found.");
                Console.WriteLine("To connect to a specific port, use ConnectToPort() method instead.");
                return false;
            }

            // Connect to the first available instance
            return ConnectToInstance(instances[0]);
        }

        /// <summary>
        /// Gets all tables in the model.
        /// </summary>
        /// <returns>A collection of tables.</returns>
        public IEnumerable<Table> GetTables()
        {
            EnsureConnected();
            return _model.Tables.Cast<Table>();
        }

        /// <summary>
        /// Gets a table by name.
        /// </summary>
        /// <param name="tableName">The name of the table.</param>
        /// <returns>The table object, or null if not found.</returns>
        public Table GetTable(string tableName)
        {
            EnsureConnected();
            return _model.Tables.Find(tableName);
        }

        /// <summary>
        /// Gets all measures in the model.
        /// </summary>
        /// <returns>A collection of measures.</returns>
        public IEnumerable<Measure> GetAllMeasures()
        {
            EnsureConnected();
            return _model.Tables.SelectMany(t => t.Measures.Cast<Measure>());
        }

        /// <summary>
        /// Creates a new measure in the specified table.
        /// </summary>
        /// <param name="tableName">The name of the table.</param>
        /// <param name="measureName">The name of the measure.</param>
        /// <param name="daxExpression">The DAX expression for the measure.</param>
        /// <param name="formatString">Optional format string for the measure.</param>
        /// <returns>The created measure.</returns>
        public Measure CreateMeasure(string tableName, string measureName, string daxExpression, string formatString = null)
        {
            EnsureConnected();
            
            var table = _model.Tables.Find(tableName);
            if (table == null)
                throw new ArgumentException($"Table '{tableName}' not found.");

            var measure = new Measure
            {
                Name = measureName,
                Expression = daxExpression
            };

            if (!string.IsNullOrEmpty(formatString))
                measure.FormatString = formatString;

            table.Measures.Add(measure);
            _model.SaveChanges();
            
            return measure;
        }

        /// <summary>
        /// Updates an existing measure.
        /// </summary>
        /// <param name="tableName">The name of the table.</param>
        /// <param name="measureName">The name of the measure.</param>
        /// <param name="daxExpression">The new DAX expression for the measure.</param>
        /// <param name="formatString">Optional new format string for the measure.</param>
        /// <returns>The updated measure.</returns>
        public Measure UpdateMeasure(string tableName, string measureName, string daxExpression, string formatString = null)
        {
            EnsureConnected();
            
            var table = _model.Tables.Find(tableName);
            if (table == null)
                throw new ArgumentException($"Table '{tableName}' not found.");

            var measure = table.Measures.Find(measureName);
            if (measure == null)
                throw new ArgumentException($"Measure '{measureName}' not found in table '{tableName}'.");

            measure.Expression = daxExpression;
            
            if (formatString != null)
                measure.FormatString = formatString;

            _model.SaveChanges();
            
            return measure;
        }

        /// <summary>
        /// Deletes a measure from the model.
        /// </summary>
        /// <param name="tableName">The name of the table.</param>
        /// <param name="measureName">The name of the measure to delete.</param>
        /// <returns>True if the measure was deleted, false otherwise.</returns>
        public bool DeleteMeasure(string tableName, string measureName)
        {
            EnsureConnected();
            
            var table = _model.Tables.Find(tableName);
            if (table == null)
                return false;

            var measure = table.Measures.Find(measureName);
            if (measure == null)
                return false;

            table.Measures.Remove(measure);
            _model.SaveChanges();
            
            return true;
        }

        /// <summary>
        /// Creates a new calculated column in the specified table.
        /// </summary>
        /// <param name="tableName">The name of the table.</param>
        /// <param name="columnName">The name of the column.</param>
        /// <param name="daxExpression">The DAX expression for the column.</param>
        /// <param name="dataType">The data type of the column.</param>
        /// <returns>The created column.</returns>
        public CalculatedColumn CreateCalculatedColumn(string tableName, string columnName, string daxExpression, DataType dataType = DataType.String)
        {
            EnsureConnected();
            
            var table = _model.Tables.Find(tableName);
            if (table == null)
                throw new ArgumentException($"Table '{tableName}' not found.");

            var column = new CalculatedColumn
            {
                Name = columnName,
                Expression = daxExpression,
                DataType = dataType
            };

            table.Columns.Add(column);
            _model.SaveChanges();
            
            return column;
        }

        /// <summary>
        /// Gets relationships in the model.
        /// </summary>
        /// <returns>A collection of relationships.</returns>
        public IEnumerable<SingleColumnRelationship> GetRelationships()
        {
            EnsureConnected();
            return _model.Relationships.Cast<SingleColumnRelationship>();
        }

        /// <summary>
        /// Creates a new relationship between two tables.
        /// </summary>
        /// <param name="fromTableName">The name of the "many" table.</param>
        /// <param name="fromColumnName">The name of the column in the "many" table.</param>
        /// <param name="toTableName">The name of the "one" table.</param>
        /// <param name="toColumnName">The name of the column in the "one" table.</param>
        /// <returns>The created relationship.</returns>
        public SingleColumnRelationship CreateRelationship(string fromTableName, string fromColumnName, string toTableName, string toColumnName)
        {
            EnsureConnected();
            
            var fromTable = _model.Tables.Find(fromTableName);
            if (fromTable == null)
                throw new ArgumentException($"Table '{fromTableName}' not found.");
                
            var toTable = _model.Tables.Find(toTableName);
            if (toTable == null)
                throw new ArgumentException($"Table '{toTableName}' not found.");
                
            var fromColumn = fromTable.Columns.Find(fromColumnName);
            if (fromColumn == null)
                throw new ArgumentException($"Column '{fromColumnName}' not found in table '{fromTableName}'.");
                
            var toColumn = toTable.Columns.Find(toColumnName);
            if (toColumn == null)
                throw new ArgumentException($"Column '{toColumnName}' not found in table '{toTableName}'.");

            var relationship = new SingleColumnRelationship
            {
                FromColumn = fromColumn,
                ToColumn = toColumn,
                FromCardinality = RelationshipEndCardinality.Many,
                ToCardinality = RelationshipEndCardinality.One
            };

            _model.Relationships.Add(relationship);
            _model.SaveChanges();
            
            return relationship;
        }

        /// <summary>
        /// Refreshes the data in the model.
        /// </summary>
        /// <param name="refreshType">The type of refresh to perform.</param>
        public void RefreshModel(RefreshType refreshType = RefreshType.Full)
        {
            EnsureConnected();
            _database.Model.RequestRefresh(refreshType);
            _database.Model.SaveChanges();
        }

        /// <summary>
        /// Gets the diagnostic log messages from the PBIProcessHelper.
        /// </summary>
        /// <returns>The diagnostic log messages.</returns>
        public IEnumerable<string> GetDiagnosticLogs()
        {
            return PBIProcessHelper.DiagnosticLog;
        }

        /// <summary>
        /// Ensures that the wrapper is connected to a Power BI Desktop model.
        /// </summary>
        private void EnsureConnected()
        {
            if (!_isConnected)
                throw new InvalidOperationException("Not connected to a Power BI Desktop model. Call ConnectToLocalPBIDesktop() first.");
        }

        /// <summary>
        /// Disposes of resources used by the wrapper.
        /// </summary>
        public void Dispose()
        {
            if (_server != null)
            {
                _server.Disconnect();
                _server.Dispose();
                _server = null;
            }
            
            _isConnected = false;
            _currentInstance = null;
        }
    }
} 