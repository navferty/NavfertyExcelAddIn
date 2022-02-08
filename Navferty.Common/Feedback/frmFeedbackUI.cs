﻿using System;
using System.Windows.Forms;

using Navferty.Common.Controls;
//using Navferty.Common.DelegatesExtensions;

using NLog;

#nullable enable

namespace Navferty.Common.Feedback
{
	internal partial class frmFeedbackUI : FormEx
	{
		private readonly ILogger logger = LogManager.GetCurrentClassLogger();

		[Obsolete("Just for Designer!", true)]
		public frmFeedbackUI() : base()
		{
			InitializeComponent();
		}

		public frmFeedbackUI(string? message = null) : base()
		{
			InitializeComponent();

			Text = Localization.UIStrings.Feedback_Title;
			lblMessage.Text = String.Format(Localization.UIStrings.Feedback_Message, FeedbackManager.MAX_USER_TEXT_LENGH);
			txtUserMessage.MaxLength = FeedbackManager.MAX_USER_TEXT_LENGH;

			if (!string.IsNullOrWhiteSpace(message))
			{
				message = message!.LimitLength(FeedbackManager.MAX_USER_TEXT_LENGH);
				txtUserMessage.Text = message;
			}


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
				  if (FeedbackManager.SendFeedbackMail(
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

			new Action(() =>
			{
				/*
				var EEE = new Exception("Test");
				throw EEE;
				 */
				FeedbackManager.ShowLogFile();
			})
				.TryCatch(true,
				Localization.UIStrings.Feedback_ErrorTitle,
				logger, "Failed to show NLog log file!");

		}
	}
}
