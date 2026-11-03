using System;
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

		private static readonly Color LinkBlue = Color.FromArgb(5, 99, 193);

		private readonly ToolTip toolTip = new ToolTip();
		private readonly string version;

		public InfoForm()
		{
			version = typeof(InfoForm).Assembly.GetName().Version?.ToString() ?? string.Empty;

			AutoScaleMode = AutoScaleMode.Dpi;
			Font = SystemFonts.MessageBoxFont;
			Text = UIStrings.Info_Title;
			FormBorderStyle = FormBorderStyle.FixedDialog;
			StartPosition = FormStartPosition.CenterScreen;
			MaximizeBox = false;
			MinimizeBox = false;
			ShowInTaskbar = false;
			AutoSize = true;
			AutoSizeMode = AutoSizeMode.GrowAndShrink;
			Padding = new Padding(20);

			var root = new TableLayoutPanel
			{
				ColumnCount = 1,
				AutoSize = true,
				AutoSizeMode = AutoSizeMode.GrowAndShrink,
				Dock = DockStyle.Fill,
			};
			root.Controls.Add(CreateHeader());
			root.Controls.Add(new Label
			{
				Text = UIStrings.Info_Description,
				MaximumSize = new Size(DpiScale(380), 0),
				AutoSize = true,
				Margin = new Padding(0, 16, 0, 12),
			});
			root.Controls.Add(CreateLink(UIStrings.Info_Website, WebsiteUrl));
			root.Controls.Add(CreateLink(UIStrings.Info_FeatureRequest, FeatureRequestUrl));
			root.Controls.Add(CreateLink(UIStrings.Info_BugReport, BugReportUrl));
			root.Controls.Add(CreateButtonRow());

			Controls.Add(root);
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing)
				toolTip.Dispose();
			base.Dispose(disposing);
		}

		private int DpiScale(int value) => (int)Math.Round(value * DeviceDpi / 96.0);

		private Control CreateHeader()
		{
			var header = new TableLayoutPanel
			{
				ColumnCount = 2,
				AutoSize = true,
				AutoSizeMode = AutoSizeMode.GrowAndShrink,
				Margin = new Padding(0),
			};

			header.Controls.Add(new PictureBox
			{
				Image = SystemIcons.Information.ToBitmap(),
				SizeMode = PictureBoxSizeMode.StretchImage,
				Size = new Size(DpiScale(40), DpiScale(40)),
				Margin = new Padding(0, 0, 12, 0),
			});

			var titles = new FlowLayoutPanel
			{
				FlowDirection = FlowDirection.TopDown,
				WrapContents = false,
				AutoSize = true,
				Margin = new Padding(0),
			};
			titles.Controls.Add(new Label
			{
				Text = UIStrings.Info_AppName,
				Font = new Font(Font.FontFamily, 12f, FontStyle.Bold),
				AutoSize = true,
				Margin = new Padding(0),
			});
			titles.Controls.Add(new Label
			{
				Text = string.Format(UIStrings.Info_Version, version),
				ForeColor = SystemColors.GrayText,
				AutoSize = true,
				Margin = new Padding(0, 2, 0, 0),
			});
			header.Controls.Add(titles);

			return header;
		}

		private Control CreateButtonRow()
		{
			var copyLink = CreateLink(UIStrings.Info_CopyVersion, null);
			copyLink.Margin = new Padding(0, 0, 0, 0);
			copyLink.LinkClicked += (s, e) =>
			{
				try
				{
					Clipboard.SetText(version);
					copyLink.Text = UIStrings.Info_Copied;
				}
				catch (System.Runtime.InteropServices.ExternalException)
				{
					// clipboard is locked by another process
				}
			};

			var ok = new Button
			{
				Text = UIStrings.OK,
				DialogResult = DialogResult.OK,
				Size = new Size(DpiScale(75), DpiScale(27)),
				Margin = new Padding(0),
			};
			AcceptButton = ok;
			CancelButton = ok;

			var row = new TableLayoutPanel
			{
				ColumnCount = 2,
				AutoSize = true,
				Dock = DockStyle.Fill,
				Margin = new Padding(0, 16, 0, 0),
			};
			row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
			row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
			copyLink.Anchor = AnchorStyles.Left;
			row.Controls.Add(copyLink);
			row.Controls.Add(ok);
			return row;
		}

		private LinkLabel CreateLink(string text, string? url)
		{
			var link = new LinkLabel
			{
				Text = text,
				AutoSize = true,
				LinkColor = LinkBlue,
				ActiveLinkColor = LinkBlue,
				Margin = new Padding(0, 0, 0, 8),
			};

			if (url != null)
			{
				toolTip.SetToolTip(link, url);
				link.LinkClicked += (s, e) =>
					Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
			}

			return link;
		}
	}
}
