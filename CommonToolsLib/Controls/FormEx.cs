using System.ComponentModel;
using System.Windows.Forms;

namespace NavfertyCommon.Controls
{
	/// <summary>Form than closes by ESC key</summary>
	public class FormEx : System.Windows.Forms.Form
	{
		public FormEx() : base()
		{
			base.KeyPreview = true;
			this.CloseByESC = true;
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		public new bool KeyPreview => true;

		/// <summary>Close form by pressing ESC key</summary>
		[DefaultValue(true)]
		public bool CloseByESC { get; set; } = true;

		protected override void OnKeyDown(KeyEventArgs e)
		{
			base.OnKeyDown(e);

			if (CloseByESC && e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.Cancel;
		}
	}
}
