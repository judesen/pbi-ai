using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using System.Text.RegularExpressions;

namespace PBITOMWrapper
{
    /// <summary>
    /// Helper class for finding Power BI Desktop processes and their associated Analysis Services instances.
    /// </summary>
    public static class PBIProcessHelper
    {
        /// <summary>
        /// Information about a running Power BI Desktop instance.
        /// </summary>
        public class PBIInstance
        {
            /// <summary>
            /// The Process ID of the Power BI Desktop instance.
            /// </summary>
            public int ProcessId { get; set; }

            /// <summary>
            /// The port number for the local Analysis Services instance.
            /// </summary>
            public int Port { get; set; }

            /// <summary>
            /// The file path of the PBIX file.
            /// </summary>
            public string FilePath { get; set; }

            /// <summary>
            /// The name of the PBIX file (without the path).
            /// </summary>
            public string FileName => System.IO.Path.GetFileName(FilePath);
        }

        /// <summary>
        /// Logs messages for debugging purposes.
        /// </summary>
        private static List<string> _logMessages = new List<string>();

        /// <summary>
        /// Enable or disable diagnostics logging (off by default).
        /// </summary>
        public static bool DiagnosticsEnabled { get; set; } = false;

        /// <summary>
        /// Get the diagnostic log messages.
        /// </summary>
        public static IEnumerable<string> DiagnosticLog => _logMessages;

        /// <summary>
        /// Clear the diagnostic log.
        /// </summary>
        public static void ClearDiagnosticLog()
        {
            _logMessages.Clear();
        }

        /// <summary>
        /// Log a diagnostic message.
        /// </summary>
        private static void LogDiagnostic(string message)
        {
            if (DiagnosticsEnabled)
            {
                _logMessages.Add($"{DateTime.Now:HH:mm:ss.fff}: {message}");
                Console.WriteLine($"[PBIProcessHelper] {message}");
            }
        }

        /// <summary>
        /// Gets all running Power BI Desktop instances and their associated Analysis Services ports.
        /// </summary>
        /// <returns>A list of Power BI Desktop instances.</returns>
        public static List<PBIInstance> GetRunningPBIInstances()
        {
            ClearDiagnosticLog();
            var result = new List<PBIInstance>();

            try
            {
                // Find all Power BI Desktop processes
                var pbiProcesses = Process.GetProcessesByName("PBIDesktop");
                LogDiagnostic($"Found {pbiProcesses.Length} PowerBI Desktop processes");

                foreach (var process in pbiProcesses)
                {
                    try
                    {
                        LogDiagnostic($"Examining PBI process {process.Id}");
                        var instance = new PBIInstance
                        {
                            ProcessId = process.Id,
                            FilePath = GetPBIFilePath(process.Id)
                        };

                        LogDiagnostic($"Found file path: {instance.FilePath}");

                        // Find the port for the Analysis Services instance associated with this PBI process
                        instance.Port = GetPBIPort(process.Id);
                        LogDiagnostic($"Detected port: {instance.Port}");

                        if (instance.Port > 0)
                        {
                            result.Add(instance);
                        }
                        else
                        {
                            LogDiagnostic("Could not determine port, trying fallback method...");
                            
                            // Fallback: try to determine port by checking active TCP connections
                            var tcpPorts = GetActiveTcpPorts();
                            LogDiagnostic($"Found {tcpPorts.Count} active TCP ports");
                            
                            // Try some common Analysis Services ports
                            foreach (var port in tcpPorts)
                            {
                                if (port >= 49152 && port <= 65535)  // Dynamic port range
                                {
                                    LogDiagnostic($"Testing TCP port {port}...");
                                    if (TestASConnection(port))
                                    {
                                        LogDiagnostic($"Port {port} appears to be an Analysis Services port");
                                        instance.Port = port;
                                        result.Add(instance);
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogDiagnostic($"Error processing PBI instance: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogDiagnostic($"Error getting Power BI instances: {ex.Message}");
            }

            if (result.Count == 0)
            {
                LogDiagnostic("No PBI instances found. Trying to scan for all Analysis Services ports...");
                
                // Last resort: try to look at all active TCP connections
                ScanForAnalysisServicesPorts(result);
            }

            return result;
        }

        /// <summary>
        /// Creates a Power BI instance with a specific port number.
        /// </summary>
        /// <param name="portNumber">The port number to use.</param>
        /// <returns>A Power BI instance using the specified port.</returns>
        public static PBIInstance CreateInstanceWithPort(int portNumber)
        {
            if (portNumber <= 0)
                throw new ArgumentException("Port number must be greater than zero.", nameof(portNumber));

            return new PBIInstance
            {
                ProcessId = -1,  // Unknown process ID
                Port = portNumber,
                FilePath = "Unknown"  // Unknown file path
            };
        }

        /// <summary>
        /// Gets the file path of the PBIX file open in the specified Power BI Desktop process.
        /// </summary>
        /// <param name="processId">The Process ID of the Power BI Desktop instance.</param>
        /// <returns>The file path of the PBIX file.</returns>
        private static string GetPBIFilePath(int processId)
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher(
                    $"SELECT CommandLine FROM Win32_Process WHERE ProcessId = {processId}"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        var commandLine = obj["CommandLine"]?.ToString() ?? "";
                        LogDiagnostic($"Process command line: {commandLine}");
                        
                        // Try to match both .pbix and .pbit files
                        var match = Regex.Match(commandLine, @"(?<="")([^""]+\.(pbix|pbit))(?="")");
                        if (match.Success)
                        {
                            return match.Value;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogDiagnostic($"Error getting file path: {ex.Message}");
            }

            return "Unknown";
        }

        /// <summary>
        /// Gets a list of active TCP ports.
        /// </summary>
        /// <returns>A list of active TCP port numbers.</returns>
        private static List<int> GetActiveTcpPorts()
        {
            var result = new List<int>();
            try
            {
                // Get all active TCP connections
                IPGlobalProperties properties = IPGlobalProperties.GetIPGlobalProperties();
                var listeners = properties.GetActiveTcpListeners();
                
                foreach (var listener in listeners)
                {
                    if (listener.Address.Equals(IPAddress.Loopback) || 
                        listener.Address.Equals(IPAddress.Any) ||
                        listener.Address.Equals(IPAddress.IPv6Loopback) ||
                        listener.Address.Equals(IPAddress.IPv6Any))
                    {
                        result.Add(listener.Port);
                    }
                }
            }
            catch (Exception ex)
            {
                LogDiagnostic($"Error getting active TCP ports: {ex.Message}");
            }
            
            return result;
        }

        /// <summary>
        /// Tests if a port is an Analysis Services port.
        /// </summary>
        /// <param name="port">The port to test.</param>
        /// <returns>True if the port appears to be an Analysis Services port, false otherwise.</returns>
        private static bool TestASConnection(int port)
        {
            try
            {
                // Very simple test - we just try to connect to the port
                using (var client = new System.Net.Sockets.TcpClient())
                {
                    // Try to connect with a short timeout
                    var result = client.BeginConnect("localhost", port, null, null);
                    var success = result.AsyncWaitHandle.WaitOne(100); // 100ms timeout
                    client.EndConnect(result);
                    
                    if (success)
                    {
                        LogDiagnostic($"Successfully connected to port {port}");
                        return true;
                    }
                }
            }
            catch
            {
                // Ignore errors
            }
            
            return false;
        }

        /// <summary>
        /// Scans for Analysis Services ports.
        /// </summary>
        /// <param name="result">The list to add found instances to.</param>
        private static void ScanForAnalysisServicesPorts(List<PBIInstance> result)
        {
            var activePorts = GetActiveTcpPorts();
            LogDiagnostic($"Scanning {activePorts.Count} active TCP ports for Analysis Services...");
            
            foreach (var port in activePorts)
            {
                // Only check ports in the dynamic range
                if (port >= 49152 && port <= 65535)
                {
                    if (TestASConnection(port))
                    {
                        LogDiagnostic($"Found potential Analysis Services port: {port}");
                        
                        // Create a new instance with just the port
                        var instance = new PBIInstance
                        {
                            ProcessId = -1,  // Unknown process ID
                            Port = port,
                            FilePath = "Unknown"  // Unknown file path
                        };
                        
                        result.Add(instance);
                    }
                }
            }
        }

        /// <summary>
        /// Gets the port number for the Analysis Services instance associated with a Power BI Desktop process.
        /// </summary>
        /// <param name="processId">The Process ID of the Power BI Desktop instance.</param>
        /// <returns>The port number, or 0 if not found.</returns>
        private static int GetPBIPort(int processId)
        {
            try
            {
                // Find all msmdsrv.exe processes (Analysis Services)
                var msmdsrvProcesses = Process.GetProcessesByName("msmdsrv");
                LogDiagnostic($"Found {msmdsrvProcesses.Length} msmdsrv processes");

                // Get all child processes of the PBI Desktop process
                using (var searcher = new ManagementObjectSearcher(
                    "SELECT * FROM Win32_Process WHERE Name = 'msmdsrv.exe'"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        // Check if this msmdsrv process was started by the Power BI Desktop process
                        var msmdsrvProcessId = Convert.ToInt32(obj["ProcessId"]);
                        var parentProcessId = GetParentProcessId(msmdsrvProcessId);
                        
                        LogDiagnostic($"msmdsrv process {msmdsrvProcessId} has parent {parentProcessId}");

                        // If this msmdsrv is a child of the Power BI Desktop process or one of its intermediaries
                        if (IsChildProcessOf(msmdsrvProcessId, processId))
                        {
                            LogDiagnostic($"msmdsrv process {msmdsrvProcessId} is a child of PBI Desktop process {processId}");
                            
                            // Get the port that this msmdsrv is listening on
                            var port = GetPortByProcessId(msmdsrvProcessId);
                            
                            if (port > 0)
                            {
                                LogDiagnostic($"Found port {port} for msmdsrv process {msmdsrvProcessId}");
                                return port;
                            }
                        }
                    }
                }
                
                // Also check the parent-child relationship in the reverse direction
                // (sometimes PBI Desktop is a child of msmdsrv)
                using (var searcher = new ManagementObjectSearcher(
                    $"SELECT * FROM Win32_Process WHERE ProcessId = {processId}"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        var parentProcessId = Convert.ToInt32(obj["ParentProcessId"]);
                        LogDiagnostic($"PBI Desktop process {processId} has parent {parentProcessId}");
                        
                        // Check if the parent process is msmdsrv.exe
                        using (var parentSearcher = new ManagementObjectSearcher(
                            $"SELECT * FROM Win32_Process WHERE ProcessId = {parentProcessId}"))
                        {
                            foreach (var parentObj in parentSearcher.Get())
                            {
                                var processName = parentObj["Name"]?.ToString();
                                LogDiagnostic($"Parent process {parentProcessId} is {processName}");
                                
                                if (processName?.ToLower() == "msmdsrv.exe")
                                {
                                    LogDiagnostic($"PBI Desktop process {processId} has msmdsrv parent {parentProcessId}");
                                    var port = GetPortByProcessId(parentProcessId);
                                    
                                    if (port > 0)
                                    {
                                        LogDiagnostic($"Found port {port} for msmdsrv process {parentProcessId}");
                                        return port;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogDiagnostic($"Error getting PBI port: {ex.Message}");
            }

            return 0;
        }

        /// <summary>
        /// Checks if a process is a direct or indirect child of another process.
        /// </summary>
        /// <param name="childProcessId">The potential child process ID.</param>
        /// <param name="potentialParentProcessId">The potential parent process ID.</param>
        /// <returns>True if childProcessId is a direct or indirect child of potentialParentProcessId.</returns>
        private static bool IsChildProcessOf(int childProcessId, int potentialParentProcessId)
        {
            try
            {
                // Check up to 5 levels of parent-child relationships
                int currentProcessId = childProcessId;
                for (int i = 0; i < 5; i++)
                {
                    int parentProcessId = GetParentProcessId(currentProcessId);
                    
                    if (parentProcessId == 0)
                        return false;
                        
                    if (parentProcessId == potentialParentProcessId)
                        return true;
                        
                    currentProcessId = parentProcessId;
                }
            }
            catch (Exception ex)
            {
                LogDiagnostic($"Error checking parent-child relationship: {ex.Message}");
            }
            
            return false;
        }

        /// <summary>
        /// Gets the parent Process ID for a given Process ID.
        /// </summary>
        /// <param name="processId">The Process ID to find the parent of.</param>
        /// <returns>The parent Process ID, or 0 if not found.</returns>
        private static int GetParentProcessId(int processId)
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher(
                    $"SELECT ParentProcessId FROM Win32_Process WHERE ProcessId = {processId}"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        return Convert.ToInt32(obj["ParentProcessId"]);
                    }
                }
            }
            catch (Exception ex)
            {
                LogDiagnostic($"Error getting parent process ID: {ex.Message}");
            }

            return 0;
        }

        /// <summary>
        /// Gets the port that a process is listening on.
        /// </summary>
        /// <param name="processId">The Process ID.</param>
        /// <returns>The port number, or 0 if not found.</returns>
        private static int GetPortByProcessId(int processId)
        {
            try
            {
                // Method 1: Using WMI to query TCP endpoints
                using (var searcher = new ManagementObjectSearcher(
                    "SELECT * FROM Win32_TCPEndpoint"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        try
                        {
                            if (obj["ProcessId"] != null)
                            {
                                var owningProcessId = Convert.ToInt32(obj["ProcessId"]);
                                if (owningProcessId == processId)
                                {
                                    var localAddress = obj["LocalAddress"]?.ToString();
                                    if (localAddress == "0.0.0.0" || localAddress == "127.0.0.1" || localAddress == "::1" || localAddress == "::")
                                    {
                                        return Convert.ToInt32(obj["LocalPort"]);
                                    }
                                }
                            }
                        }
                        catch
                        {
                            // Skip this endpoint if we can't get its information
                        }
                    }
                }

                // Method 2: Use .NET to get the TCP connections
                IPGlobalProperties properties = IPGlobalProperties.GetIPGlobalProperties();
                var connections = properties.GetActiveTcpConnections();
                
                // Get the process's active connections
                var processConnections = connections.Where(c => 
                    (c.LocalEndPoint.Address.Equals(IPAddress.Loopback) || 
                     c.LocalEndPoint.Address.Equals(IPAddress.Any) || 
                     c.LocalEndPoint.Address.Equals(IPAddress.IPv6Loopback) || 
                     c.LocalEndPoint.Address.Equals(IPAddress.IPv6Any)) &&
                    CheckProcessIdForPort(processId, c.LocalEndPoint.Port));
                
                foreach (var conn in processConnections)
                {
                    return conn.LocalEndPoint.Port;
                }
                
                // Method 3: Look at TCP listeners
                var listeners = properties.GetActiveTcpListeners();
                foreach (var listener in listeners)
                {
                    if (listener.Address.Equals(IPAddress.Loopback) || 
                        listener.Address.Equals(IPAddress.Any) || 
                        listener.Address.Equals(IPAddress.IPv6Loopback) || 
                        listener.Address.Equals(IPAddress.IPv6Any))
                    {
                        if (CheckProcessIdForPort(processId, listener.Port))
                        {
                            return listener.Port;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogDiagnostic($"Error getting port by process ID: {ex.Message}");
            }

            return 0;
        }

        /// <summary>
        /// Checks if a process is using a specific port.
        /// </summary>
        /// <param name="processId">The Process ID.</param>
        /// <param name="port">The port to check.</param>
        /// <returns>True if the process is using the port, false otherwise.</returns>
        private static bool CheckProcessIdForPort(int processId, int port)
        {
            try
            {
                // This is a simplified implementation - a real implementation would use
                // the Windows API to check if the process is using the port
                return TestASConnection(port);
            }
            catch
            {
                return false;
            }
        }
    }
} 