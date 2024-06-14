
using Navferty.Common;

#nullable enable

namespace NavfertyExcelAddIn.Feedback
{
	public class FeedbackBuilder : IFeedback
	{
		internal readonly IDialogService dialogService;
		private Microsoft.Office.Interop.Excel.Application App => Globals.ThisAddIn.Application;

		public FeedbackBuilder(IDialogService dialogService)
			=> this.dialogService = dialogService;


		public void DisplayFeedbackUI()
		{
			Navferty.Common.Feedback.FeedbackManager.ShowFeedbackUI();
		}
	}
}
