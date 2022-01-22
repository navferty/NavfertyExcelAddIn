using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NavfertyExcelAddIn.Commons.Controls
{
	/// <summary>Form than closes by ESC key</summary>
	internal class FormEx : System.Windows.Forms.Form
	{
		public FormEx() : base()
		{
			base.KeyPreview = true;
			this.KeyDown += (s, e) => { if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.Cancel; };
		}

		public new bool KeyPreview => true;
	}
}
