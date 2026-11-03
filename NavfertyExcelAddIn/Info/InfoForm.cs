using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.Info
{
	internal sealed class InfoForm : Form
	{
		private const string WebsiteUrl = "https://navferty.net";
		private const string FeatureRequestUrl = "https://github.com/navferty/NavfertyExcelAddIn/issues/new?labels=enhancement";
		private const string BugReportUrl = "https://github.com/navferty/NavfertyExcelAddIn/issues/new?labels=bug";

		public InfoForm()
		{
			Text = UIStrings.ResourceManager.GetString("Info_Title");
			FormBorderStyle = FormBorderStyle.FixedDialog;
			StartPosition = FormStartPosition.CenterScreen;
			MaximizeBox = false;
			MinimizeBox = false;
			ShowInTaskbar = false;
			AutoSize = true;
			AutoSizeMode = AutoSizeMode.GrowAndShrink;
			Padding = new Padding(12);

			var version = typeof(InfoForm).Assembly.GetName().Version;
			var panel = new FlowLayoutPanel
			{
				FlowDirection = FlowDirection.TopDown,
				WrapContents = false,
				AutoSize = true,
				AutoSizeMode = AutoSizeMode.GrowAndShrink,
			};

			panel.Controls.Add(new Label
			{
				Text = UIStrings.ResourceManager.GetString("Info_Description"),
				MaximumSize = new Size(380, 0),
				AutoSize = true,
				Margin = new Padding(3, 3, 3, 8),
			});
			panel.Controls.Add(new Label
			{
				Text = string.Format(UIStrings.ResourceManager.GetString("Info_Version") ?? "{0}", version),
				AutoSize = true,
				Margin = new Padding(3, 3, 3, 8),
			});
			panel.Controls.Add(CreateLink("Info_Website", WebsiteUrl));
			panel.Controls.Add(CreateLink("Info_FeatureRequest", FeatureRequestUrl));
			panel.Controls.Add(CreateLink("Info_BugReport", BugReportUrl));

			var ok = new Button
			{
				Text = UIStrings.OK,
				DialogResult = DialogResult.OK,
				Margin = new Padding(3, 12, 3, 3),
			};
			panel.Controls.Add(ok);

			AcceptButton = ok;
			CancelButton = ok;
			Controls.Add(panel);
		}

		private static LinkLabel CreateLink(string textKey, string url)
		{
			var link = new LinkLabel
			{
				Text = UIStrings.ResourceManager.GetString(textKey),
				AutoSize = true,
			};
			link.LinkClicked += (s, e) =>
				Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
			return link;
		}
	}
}
