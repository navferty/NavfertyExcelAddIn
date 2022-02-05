using System.Collections.Generic;

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
