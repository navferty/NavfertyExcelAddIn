
using System.Linq;

using NLog;
using NLog.Layouts;
using NLog.Targets;

#nullable enable

namespace Navferty.Common
{
	internal static class LogManagement
	{
		//private const string DefaultTargetName = "AllTargets";

		public static string? GetTargetFilename(string? targetName)
		{
			FileTarget? target = null;
			if (string.IsNullOrWhiteSpace(targetName))
			{
				var fileTargets = GetFileTargets();
				target = fileTargets.FirstOrDefault();
			}
			else
			{
				target = GetTarget<FileTarget>(targetName!);
			}
			if (null == target) return null;

			var layout = target.FileName as SimpleLayout;
			if (null == layout) return null;

			// layout.Text provides the filename "template"
			// LogEventInfo is required; might make sense for a log line template but not really filename
			//var filename = layout.Render(new LogEventInfo()).Replace(@"/", @"");
			var filename = layout.Render(new LogEventInfo()).Replace(@"/", @"\").Replace(@"\\", @"\"); ;
			return filename;
		}

		private static T? GetTarget<T>(string targetName)
			where T : Target
		{
			if (null == LogManager.Configuration) return null;
			var target = LogManager.Configuration.FindTargetByName(targetName) as T;
			return target;
		}

		private static T[]? GetTargets<T>()
			where T : Target
		{
			if (null == LogManager.Configuration) return null;

			var targets = LogManager.Configuration.AllTargets?.Where(trg => trg.GetType() == typeof(T)).ToArray();
			return (T[]?)targets;
		}
		private static FileTarget[]? GetFileTargets()
		{
			if (null == LogManager.Configuration) return null;

			var targets = LogManager.Configuration.AllTargets?
			.Select(trg =>
			{
				FileTarget? ft = null;
				if (trg is FileTarget ft2) ft = ft2;
				return ft;
			})
			.Where(ft => ft != null)
			.Select(ft => (FileTarget)ft!)
			.ToArray();
			return targets;
		}
	}
}
