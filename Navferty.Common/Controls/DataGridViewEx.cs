using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

#nullable enable

namespace Navferty.Common.Controls
{
	/// <summary>DataGridView that do not blick return key</summary>
	public class DataGridViewEx : System.Windows.Forms.DataGridView
	{

		protected string emptyText = string.Empty;
		protected ContentAlignment emptyTextAlign = ContentAlignment.MiddleCenter;

		public DataGridViewEx() : base()
		{
			SetStyle(
				ControlStyles.ResizeRedraw
				| ControlStyles.AllPaintingInWmPaint
				| ControlStyles.DoubleBuffer
				| ControlStyles.OptimizedDoubleBuffer
				, true);
		}

		/// <summary>Like StandardTab but for the Enter key.</summary>
		[Category("Behavior"), Description("Disable default edit/advance to next row behavior of of the Enter key.")]
		public bool StandardEnter { get; set; }

		/// <summary>Implement StandardEnter.</summary>
		protected override bool IsInputKey(Keys keyData)
		{
			if (StandardEnter && keyData == Keys.Enter)
				// Force this key to be treated like something to pass
				// to ProcessDialogKey() (like the Enter key normally
				// would be for controls which aren’t DataGridView).
				return false;

			return base.IsInputKey(keyData);
		}

		private static readonly MethodInfo _Control_ProcessDialogKey = typeof(Control).GetMethod("ProcessDialogKey", BindingFlags.Instance | BindingFlags.NonPublic);

		protected override bool ProcessDialogKey(Keys keyData)
		{
			if (StandardEnter && keyData == Keys.Enter)
				// Copy the default implementation of
				// Control.ProcessDialogKey(). Since we can’t access
				// the base class (DataGridView)’s base class’s
				// implementation directly, and since we cannot
				// legally access Control.ProcessDialogKey() on other
				// Control object, we are forced to use reflection.
				return Parent == null ? false : (bool)_Control_ProcessDialogKey.Invoke(Parent, new object[] { keyData, });

			return base.ProcessDialogKey(keyData);
		}

		[Localizable(true)]
		[DefaultValue("")]
		[Description("Text than displayed when ListBox does not contains any items")]
		public string EmptyText
		{
			get => emptyText; set { emptyText = value; Invalidate(); }
		}

		[DefaultValue(ContentAlignment.MiddleCenter)]
		public virtual ContentAlignment EmptyTextAlign
		{
			get => emptyTextAlign;
			set
			{
				if (value == emptyTextAlign) return;
				emptyTextAlign = value;
				Invalidate();
			}
		}

		protected override void OnPaintBackground(PaintEventArgs e)
		{
			base.OnPaintBackground(e);

			Graphics g = e.Graphics;
			g.Clear(BackgroundColor);
		}
		protected override void OnPaint(PaintEventArgs e)
		{
			base.OnPaint(e);
			if ((RowCount > 0) || (string.IsNullOrWhiteSpace(emptyText))) return;

			Graphics g = e.Graphics;
			g.PageUnit = GraphicsUnit.Pixel;

			g.SetClip(ClientRectangle);
			g.Clear(BackgroundColor);

			g.DrawTextEx(emptyText, Font, ForeColor, ClientRectangle, emptyTextAlign);
		}
	}
}
