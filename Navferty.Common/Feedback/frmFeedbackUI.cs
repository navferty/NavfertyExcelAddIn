using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Navferty.Common.Controls;

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
			lblSummary.Text = Localization.UIStrings.Feedback_Summary;

			chkIncludeScreenshots.Text = Localization.UIStrings.Feedback_IncludeScreenshots;
			chkIncludeScreenshots.Checked = true;

			llGotoGithub.Text = Localization.UIStrings.Feedback_GotoGithub;
			btnSend.Text = Localization.UIStrings.Feedback_Send;
		}

		private void OnSend(object sender, EventArgs e)
		{
			try
			{
				if (FeedbackManager.SendFeedEMail(
					txtUserMessage.Text.Trim(),
					chkIncludeScreenshots.Checked,
					this))

					DialogResult = DialogResult.OK;
			}
			catch (Exception ex)
			{
				logger.Error(ex, "Failed to send feedback email!");
				MessageBox.Show(ex.Message,
					Localization.UIStrings.Feedback_ErrorTitle,
					MessageBoxButtons.OK,
					MessageBoxIcon.Error);
			}
		}
	}
}
