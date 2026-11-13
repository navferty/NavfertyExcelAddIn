using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;

internal static class Program
{
    private static int Main(string[] args)
    {
        if (args.Length > 1 || (args.Length == 1 && args[0] != "--no-browser"))
        {
            Console.Error.WriteLine("Usage: Website [--no-browser]");
            return 2;
        }

        try
        {
            var source = FindWebsiteDirectory();
            using (var job = new ChildProcessJob())
            using (var server = new Process())
            using (var ready = new ManualResetEventSlim())
            {
                string? address = null;
                server.StartInfo = new ProcessStartInfo
                {
                    FileName = "hugo",
                    Arguments = "server --bind 127.0.0.1 --baseURL http://localhost --port 1313",
                    WorkingDirectory = source,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };
                server.OutputDataReceived += (_, e) =>
                {
                    if (e.Data == null) return;
                    Console.WriteLine(e.Data);
                    if (e.Data.StartsWith("hugo v", StringComparison.OrdinalIgnoreCase) &&
                        !e.Data.StartsWith("hugo v0.123.7+extended", StringComparison.OrdinalIgnoreCase))
                        Console.Error.WriteLine("Warning: production uses Hugo Extended 0.123.7; this version may produce different output.");
                    var match = Regex.Match(e.Data, @"Web Server is available at (http://localhost:\d+/)");
                    if (match.Success)
                    {
                        address = match.Groups[1].Value;
                        ready.Set();
                    }
                };
                server.ErrorDataReceived += (_, e) =>
                {
                    if (e.Data != null) Console.Error.WriteLine(e.Data);
                };

                try
                {
                    server.Start();
                }
                catch (Win32Exception error)
                {
                    throw new InvalidOperationException(
                        "Could not start Hugo. Install Hugo Extended 0.123.7 and ensure hugo is on PATH.", error);
                }

                try
                {
                    job.Add(server);
                    server.BeginOutputReadLine();
                    server.BeginErrorReadLine();
                    Console.CancelKeyPress += (_, e) =>
                    {
                        e.Cancel = true;
                        if (!server.HasExited) server.Kill();
                    };

                    var timeout = Stopwatch.StartNew();
                    while (!ready.Wait(100))
                    {
                        if (server.HasExited)
                        {
                            server.WaitForExit();
                            Console.Error.WriteLine("Hugo exited before the website was ready.");
                            return server.ExitCode == 0 ? 1 : server.ExitCode;
                        }
                        if (timeout.Elapsed > TimeSpan.FromSeconds(60))
                            throw new TimeoutException("Hugo did not report a ready server within 60 seconds.");
                    }

                    if (address == null)
                        throw new InvalidOperationException("Hugo did not provide a server address.");
                    if (args.Length == 0)
                    {
                        try
                        {
                            Process.Start(new ProcessStartInfo(address) { UseShellExecute = true });
                        }
                        catch (Exception error) when (error is Win32Exception || error is InvalidOperationException)
                        {
                            Console.Error.WriteLine("Could not open the browser: " + error.Message);
                            Console.Error.WriteLine("Open " + address + " manually.");
                        }
                    }
                    Console.WriteLine("Press Ctrl+C or stop debugging to stop Hugo.");
                    server.WaitForExit();
                    return server.ExitCode;
                }
                finally
                {
                    if (!server.HasExited)
                    {
                        server.Kill();
                        server.WaitForExit();
                    }
                }
            }
        }
        catch (Exception error) when (error is InvalidOperationException || error is IOException ||
                                      error is Win32Exception || error is TimeoutException)
        {
            Console.Error.WriteLine(error.Message);
            return 1;
        }
    }

    private static string FindWebsiteDirectory()
    {
        var directory = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
        while (directory != null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "hugo.toml")))
                return directory.FullName;
            directory = directory.Parent;
        }
        throw new DirectoryNotFoundException("Cannot find Website/hugo.toml above the launcher executable.");
    }
}
