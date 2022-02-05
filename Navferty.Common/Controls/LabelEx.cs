using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

#nullable enable

namespace Navferty.Common.Controls
{
	public class LabelEx : System.Windows.Forms.Label
	{
		public LabelEx() : base()
		{
			SetStyle(
				ControlStyles.UserPaint
				| ControlStyles.Opaque
				| ControlStyles.SupportsTransparentBackColor
				| ControlStyles.ResizeRedraw
				| ControlStyles.DoubleBuffer
				| ControlStyles.OptimizedDoubleBuffer
				| ControlStyles.AllPaintingInWmPaint
				| ControlStyles.CacheText
				, true);
		}


		protected override void OnPaintBackground(PaintEventArgs e)
		{
			base.OnPaintBackground(e);
		}

		private bool drawAsInfotip = true;
		[DefaultValue(true)]
		public bool DrawAsInfotip
		{
			get => drawAsInfotip;
			set
			{
				drawAsInfotip = value;
				Invalidate();
			}
		}

		private const int DEFAULT_infotipCornersRadius = 8 * 2;
		private int infotipCornersRadius = DEFAULT_infotipCornersRadius;
		[DefaultValue(DEFAULT_infotipCornersRadius)]
		public int InfotipCornersRadius
		{
			get => infotipCornersRadius;
			set
			{
				infotipCornersRadius = value;
				Invalidate();
			}
		}

		private Color infotipBorderColor = SystemColors.ButtonShadow;
		[DefaultValue(typeof(SystemColors), "ButtonShadow")]
		public Color InfotippBorderColor
		{
			get => infotipBorderColor;
			set
			{
				infotipBorderColor = value;
				Invalidate();
			}
		}

		protected override void OnPaint(PaintEventArgs e)
		{
			OnPaintBackground(e);
			//base.OnPaint(e);
			if (string.IsNullOrWhiteSpace(Text)) return;

			Graphics g = e.Graphics;
			g.PageUnit = GraphicsUnit.Pixel;
			g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;

			g.Clear(BackColor);

			Color clrText = ForeColor;
			Rectangle rcInfo = ClientRectangle;
			if (drawAsInfotip)
			{
				rcInfo.Width -= 1;
				rcInfo.Height -= 1;

				rcInfo.DeflateByPadding(Padding);

				using (var pnFrame = new Pen(infotipBorderColor, 1))
					g.DrawRoundRect(rcInfo, infotipCornersRadius, pnFrame, SystemBrushes.Info);

				clrText = SystemColors.InfoText;
			}
			g.DrawTextEx(Text, Font, clrText, rcInfo, TextAlign, StringFormatFlags.NoWrap);
		}
	}
}
