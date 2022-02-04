using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

#nullable enable

namespace NavfertyCommon
{

	namespace Controls
	{
		/// <summary>Used for draw custom 'EmptyText' when list does not contains items</summary>
		public class CheckedListBoxEx : System.Windows.Forms.CheckedListBox
		{
			protected string emptyText = string.Empty;
			protected ContentAlignment emptyTextAlign = ContentAlignment.MiddleCenter;

			public CheckedListBoxEx() : base() { }

			#region Properties

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

			#endregion

			#region Custom Paint

			/// <summary>We must to override WndProc bc Standart CheckedListBox does not allow to Paint over itself</summary>
			protected override void WndProc(ref Message m)
			{
				//First doing default message processing...
				base.WndProc(ref m);

				//Now do our job.
				switch ((WinAPI.WindowMessages)m.Msg)
				{
					case WinAPI.WindowMessages.WM_PAINT: RePaint(); break;
				}
			}

			private void RePaint()
			{
				if (Items.Count > 0 || string.IsNullOrWhiteSpace(emptyText))
					return;//We paint over only if ListBox noes not have any items and and EmptyText

				var rcClient = WinAPI.GetClientRect(this);
				using (var dc = new WinAPI.DC(this))
				using (var g = dc.CreateGraphics())
				{
					g.PageUnit = GraphicsUnit.Pixel;
					using (var sf = emptyTextAlign.ToStringFormat())
					using (var brText = new SolidBrush(SystemColors.ControlDarkDark))
						g.DrawString(emptyText, Font, brText, rcClient, sf);

				}
			}

			#endregion
		}
	}

}
