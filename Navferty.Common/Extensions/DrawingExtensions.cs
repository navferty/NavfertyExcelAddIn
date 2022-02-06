using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

#nullable enable

namespace Navferty.Common
{
	[DebuggerStepThrough]
	public static class DrawingExtensions
	{

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		internal static StringAlignment GetAlignment(this ContentAlignment ca)
			=> ca switch
			{
				var ca2
				when ca2 == ContentAlignment.TopLeft || ca2 == ContentAlignment.MiddleLeft || ca2 == ContentAlignment.BottomLeft
				=> StringAlignment.Near,

				var ca2
				when ca2 == ContentAlignment.TopCenter || ca2 == ContentAlignment.MiddleCenter || ca2 == ContentAlignment.BottomCenter
				=> StringAlignment.Center,

				var ca2
				when ca2 == ContentAlignment.TopRight || ca2 == ContentAlignment.MiddleRight || ca2 == ContentAlignment.BottomRight
					=> StringAlignment.Far,

				_ => StringAlignment.Center,
			};


		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		internal static StringAlignment GetLineAlignment(this ContentAlignment ca)
			=> ca switch
			{
				var ca2
				when ca2 == ContentAlignment.TopLeft || ca2 == ContentAlignment.TopCenter || ca2 == ContentAlignment.TopRight
				=> StringAlignment.Near,

				var ca2
				when ca2 == ContentAlignment.MiddleLeft || ca2 == ContentAlignment.MiddleCenter || ca2 == ContentAlignment.MiddleRight
				=> StringAlignment.Center,

				var ca2
				when ca2 == ContentAlignment.BottomLeft || ca2 == ContentAlignment.BottomCenter || ca2 == ContentAlignment.BottomRight
					=> StringAlignment.Far,

				_ => StringAlignment.Center,
			};


		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		internal static StringFormat ToStringFormat(this ContentAlignment ca)
		{
			var sf = new StringFormat()
			{
				Alignment = ca.GetAlignment(),
				LineAlignment = ca.GetLineAlignment()
			};
			return sf;
		}


		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		internal static void DrawTextEx(
			this Graphics g,
			string text,
			Font font,
			Color textcolor,
			Rectangle rect,
			ContentAlignment textAlign,
			StringFormatFlags? additionalFF = null)
		{
			using (Brush brText = new SolidBrush(textcolor))
				g.DrawTextEx(text, font, brText, rect, textAlign, additionalFF);
		}


		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		internal static void DrawTextEx(
			this Graphics g,
			string text,
			Font font,
			Brush textbrush,
			Rectangle rect,
			ContentAlignment textAlign,
			StringFormatFlags? additionalFF = null)
		{
			using (var sf = textAlign.ToStringFormat())
			{
				if (additionalFF != null && additionalFF.HasValue) sf.FormatFlags |= additionalFF.Value;
				g.DrawString(text, font, textbrush, rect, sf);
			}
		}


		[DebuggerNonUserCode, DebuggerStepThrough, MethodImpl(MethodImplOptions.AggressiveInlining)]
		internal static GraphicsPath CreateRoundRect(this Rectangle rc, int radius)
		{
			Rectangle rcCorner = new(rc.X, rc.Y, radius, radius);

			GraphicsPath path = new();
			path.AddArc(rcCorner, 180, 90);

			rcCorner.X = rc.X + rc.Width - radius;
			path.AddArc(rcCorner, 270, 90);

			rcCorner.Y = rc.Y + rc.Height - radius;
			path.AddArc(rcCorner, 0, 90);

			rcCorner.X = rc.X;
			path.AddArc(rcCorner, 90, 90);

			path.CloseFigure();
			return path;
		}


		[DebuggerNonUserCode, DebuggerStepThrough, MethodImpl(MethodImplOptions.AggressiveInlining)]
		internal static void DrawPath(
			   this Graphics g,
			   GraphicsPath path,
			   Pen? pen = null,
			   Brush? brush = null)
		{
			if (brush != null) g.FillPath(brush, path);
			if (pen != null) g.DrawPath(pen, path);
		}


		[DebuggerNonUserCode, DebuggerStepThrough, MethodImpl(MethodImplOptions.AggressiveInlining)]
		internal static void DrawRoundRect(
			   this Graphics g,
			   Rectangle rc,
			   int radius,
			   Pen? pen = null,
			   Brush? brush = null)
		{
			using (var path = rc.CreateRoundRect(radius))
				g.DrawPath(path, pen, brush);
		}


		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		internal static void DrawPathShadow(this Graphics g, GraphicsPath gp, int radius, int intensity = 100)
		{
			double alpha = 0;
			double astep = 0;
			double astepstep = (double)intensity / radius / (radius / 2D);
			for (int thickness = radius; thickness > 0; thickness--)
			{
				using (Pen p = new Pen(Color.FromArgb((int)alpha, 0, 0, 0), thickness))
				{
					p.LineJoin = LineJoin.Round;
					g.DrawPath(p, gp);
				}
				alpha += astep;
				astep += astepstep;
			}
		}


		[DebuggerNonUserCode, DebuggerStepThrough, MethodImpl(MethodImplOptions.AggressiveInlining)]
		internal static void DeflateByPadding(
			   this ref Rectangle rect,
			   Padding pd)
		{
			if (null == pd) return;

			if (pd.Left > 0)
			{
				rect.Offset(pd.Left, 0);
				rect.Width -= pd.Left;
			}
			if (pd.Right > 0) rect.Width -= pd.Right;

			if (pd.Top > 0)
			{
				rect.Offset(0, pd.Top);
				rect.Height -= pd.Top;
			}
			if (pd.Bottom > 0) rect.Height -= pd.Bottom;
		}


	}
}
