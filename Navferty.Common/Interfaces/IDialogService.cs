using System.Collections.Generic;

#nullable enable

namespace Navferty.Common
{
	public interface IDialogService
	{
		void ShowError(string message);
		void ShowInfo(string message);
		bool Ask(string message, string caption);
		void ShowVersion();

		IReadOnlyCollection<string> AskForFiles(bool allowMultiselect, FileType fileType);
		string AskFileNameSaveAs(string initialFileName, FileType fileType);
		string AskForSaveFile(FileType fileType);
	}

	public enum FileType
	{
		All = 0,
		Excel = 1,
		Xml = 2,
		Xsd = 3,
		Pdf = 4,
		Db = 5
	}
}
