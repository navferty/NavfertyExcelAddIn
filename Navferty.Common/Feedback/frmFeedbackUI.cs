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
		}

		private void button1_Click(object sender, EventArgs e)
		{
			try
			{
				if (FeedbackManager.SendFeedEMail("Test message", true))
					this.DialogResult = DialogResult.OK;
			}
			catch (Exception ex)
			{
				logger.Error(ex, "Failed to send feedback email!");
				MessageBox.Show(ex.Message, "Failed to send feedback email!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}
