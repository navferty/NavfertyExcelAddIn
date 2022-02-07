using System;
using System.Windows.Forms;

using Navferty.Common.Controls;
//using Navferty.Common.DelegatesExtensions;

using NLog;

namespace Navferty.Common.Feedback
{
	internal partial class frmFeedbackUI : FormEx
	{
		private readonly ILogger logger = LogManager.GetCurrentClassLogger();

		public frmFeedbackUI()
		{
			InitializeComponent();

			Text = Localization.UIStrings.Feedback_Title;
			lblMessage.Text = String.Format(Localization.UIStrings.Feedback_Message, FeedbackManager.MAX_USER_TEXT_LENGH);
			txtUserMessage.MaxLength = FeedbackManager.MAX_USER_TEXT_LENGH;

			chkIncludeScreenshots.Text = Localization.UIStrings.Feedback_IncludeScreenshots;
			chkIncludeScreenshots.Checked = true;

			llGotoGithub.Text = Localization.UIStrings.Feedback_GotoGithub;
			btnSend.Text = Localization.UIStrings.Feedback_Send;

			//Create ink to show NLog Log file
			{
				string fullSummaryText = string.Format(Localization.UIStrings.Feedback_Summary_Template, Localization.UIStrings.Feedback_Summary_Loglink);
				lblSummary.Text = fullSummaryText;
				var la = new LinkArea(fullSummaryText.IndexOf(Localization.UIStrings.Feedback_Summary_Loglink), Localization.UIStrings.Feedback_Summary_Loglink.Length);
				lblSummary.LinkArea = la;
				lblSummary.LinkClicked += (s, e) => OnShowLog();
			}
		}

		private void OnSend(object sender, EventArgs e)
		{
			new Action(() =>
			  {
				  if (FeedbackManager.SendFeedEMail(
					  txtUserMessage.Text.Trim(),
					  chkIncludeScreenshots.Checked,
					  this))

					  DialogResult = DialogResult.OK;

			  }).TryCatch(true,
				Localization.UIStrings.Feedback_ErrorTitle,
				logger, "Failed to send feedback mail!");
		}


		private void OnGotoGithub(object sender, LinkLabelLinkClickedEventArgs e)
		{
			new Action(() => { FeedbackManager.ShowGithub(); })
				.TryCatch(true,
				Localization.UIStrings.Feedback_ErrorTitle,
				logger, "Failed to open Github bugtracker!");
		}
		private void OnShowLog()
		{
			new Action(() => { FeedbackManager.ShowLogFile(); })
				.TryCatch(true,
				Localization.UIStrings.Feedback_ErrorTitle,
				logger, "Failed to show NLog log file!");

		}
	}
}
