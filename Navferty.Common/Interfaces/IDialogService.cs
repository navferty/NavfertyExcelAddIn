using System;
using System.Collections.Generic;

using NLog;

#nullable enable

namespace Navferty.Common
{
	public interface IDialogService
	{
		[Obsolete("Use ShowError(Exception, ILogger) instead", true)]
		void ShowError(string message);
		void ShowError(Exception e, ILogger? logger = null);
		void ShowInfo(string message);
		bool Ask(string message, string caption);
		void ShowVersion();

		IReadOnlyCollection<string> AskForFiles(bool allowMultiselect, FileType fileType);
		string AskFileNameSaveAs(string initialFileName, FileType fileType);
	}

	public enum FileType
	{
		All = 0,
		Excel = 1,
		Xml = 2,
		Xsd = 3,
		Pdf = 4
	}
}
