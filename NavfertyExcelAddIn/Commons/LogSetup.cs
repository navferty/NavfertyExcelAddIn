using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Xml;

using NLog;
using NLog.Config;

namespace NavfertyExcelAddIn.Commons
{
	/// <summary>
	/// Loads the embedded NLog configuration, manages the log level and logs unhandled exceptions.
	/// </summary>
	internal static class LogSetup
	{
		private const string ConfigResourceName = "NavfertyExcelAddIn.NLog.config";
		private const string LogDirVariable = "logDir";
		private const int LogRetentionDays = 14;

		private static readonly object SyncRoot = new object();
		private static bool initialized;
		private static bool verbose;

		public static string LogDirectory { get; private set; } = string.Empty;

		public static bool IsVerbose
		{
			get => verbose;
			set
			{
				verbose = value;
				ApplyLevel();
				LogManager.GetLogger(nameof(LogSetup)).Info($"Verbose (Debug) logging {(value ? "enabled" : "disabled")}");
			}
		}

		public static void Initialize()
		{
			lock (SyncRoot)
			{
				if (initialized)
					return;
				initialized = true;
			}

			using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(ConfigResourceName))
			using (var reader = XmlReader.Create(stream))
			{
				var config = new XmlLoggingConfiguration(reader, null);
				LogManager.Configuration = config;
				LogDirectory = config.Variables[LogDirVariable].Render(LogEventInfo.CreateNullEvent());
			}

			DeleteOldLogs();

			AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
			TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
		}

		public static void Shutdown()
		{
			AppDomain.CurrentDomain.UnhandledException -= OnUnhandledException;
			TaskScheduler.UnobservedTaskException -= OnUnobservedTaskException;
			LogManager.Shutdown();
		}

		private static void ApplyLevel()
		{
			var config = LogManager.Configuration;
			if (config == null)
				return;

			var minLevel = verbose ? LogLevel.Debug : LogLevel.Info;
			foreach (var rule in config.LoggingRules)
			{
				rule.SetLoggingLevels(minLevel, LogLevel.Fatal);
			}

			LogManager.ReconfigExistingLoggers();
		}

		private static void DeleteOldLogs()
		{
			try
			{
				if (!Directory.Exists(LogDirectory))
					return;

				var threshold = DateTime.UtcNow.AddDays(-LogRetentionDays);
				foreach (var file in Directory.GetFiles(LogDirectory, "*.log")
					.Where(f => File.GetLastWriteTimeUtc(f) < threshold))
				{
					File.Delete(file);
				}
			}
			catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
			{
				LogManager.GetLogger(nameof(LogSetup)).Warn(ex, "Failed to delete old log files");
			}
		}

		// The AppDomain is shared with Excel and other add-ins, so the exception is not necessarily ours.
		private static void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
		{
			var logger = LogManager.GetLogger("UnhandledException");
			logger.Fatal(e.ExceptionObject as Exception, $"Unhandled exception, terminating: {e.IsTerminating}");
			LogManager.Flush(TimeSpan.FromSeconds(3));
		}

		private static void OnUnobservedTaskException(object sender, UnobservedTaskExceptionEventArgs e)
		{
			var logger = LogManager.GetLogger("UnobservedTaskException");
			logger.Error(e.Exception, "Unobserved task exception");
		}
	}
}
