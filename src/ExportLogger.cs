using System;
using System.IO;
using System.Text;

namespace BimAutomationTool
{
    public class ExportLogger : IDisposable
    {
        private StreamWriter _writer;
        private string _logPath;
        private bool _disposed = false;

        public ExportLogger(string logPath)
        {
            _logPath = logPath;
            try
            {
                _writer = new StreamWriter(logPath, false, Encoding.UTF8);
                _writer.AutoFlush = true;
                LogInfo("=== BIM Automation Tool - Export Log ===");
                LogInfo($"Started: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                LogInfo("");
            }
            catch (Exception ex)
            {
                // If we can't create log file, continue without logging
                Console.WriteLine($"Failed to create log file: {ex.Message}");
            }
        }

        public void LogInfo(string message)
        {
            Log("INFO", message);
        }

        public void LogWarning(string message)
        {
            Log("WARN", message);
        }

        public void LogError(string message)
        {
            Log("ERROR", message);
        }

        private void Log(string level, string message)
        {
            if (_writer == null) return;

            try
            {
                string timestamp = DateTime.Now.ToString("HH:mm:ss");
                _writer.WriteLine($"[{timestamp}] [{level}] {message}");
            }
            catch
            {
                // Ignore logging errors
            }
        }

        public void Close()
        {
            if (_writer != null && !_disposed)
            {
                try
                {
                    LogInfo("");
                    LogInfo($"Completed: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    LogInfo("=== End of Log ===");
                    _writer.Close();
                }
                catch { }
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                Close();
                _disposed = true;
            }
        }
    }
}
