using System;
using System.Diagnostics;
using System.IO;

namespace LksWrapper
{
    internal static class Program
    {
        private static int Main(string[] args)
        {
            var baseDir = AppContext.BaseDirectory;
            var pythonScript = Path.Combine(baseDir, "LKS.py");

            if (!File.Exists(pythonScript))
            {
                Console.Error.WriteLine($"Python script not found: {pythonScript}");
                return 1;
            }

            var pythonExe = FindPythonExecutable();
            if (pythonExe == null)
            {
                Console.Error.WriteLine("Python runtime not found. Please install Python and ensure it is on PATH.");
                return 1;
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = pythonExe,
                Arguments = $"\"{pythonScript}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = false,
                WorkingDirectory = baseDir
            };

            using var process = new Process { StartInfo = startInfo };
            process.OutputDataReceived += (_, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    Console.WriteLine(e.Data);
                }
            };
            process.ErrorDataReceived += (_, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    Console.Error.WriteLine(e.Data);
                }
            };

            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            process.WaitForExit();

            return process.ExitCode;
        }

        private static string? FindPythonExecutable()
        {
            var candidates = new[] { "python", "python3" };
            foreach (var candidate in candidates)
            {
                if (CanStart(candidate, "--version"))
                {
                    return candidate;
                }
            }

            return null;
        }

        private static bool CanStart(string fileName, string arguments)
        {
            try
            {
                using var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = fileName,
                        Arguments = arguments,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };

                process.Start();
                process.WaitForExit(3000);
                return process.ExitCode == 0;
            }
            catch
            {
                return false;
            }
        }
    }
}
