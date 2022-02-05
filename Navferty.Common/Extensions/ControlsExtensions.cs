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
	public static class ControlsExtensions
	{
		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static void SetVistaCueBanner(this TextBox ctl, string? BannerText = null)
		{
			_ = ctl ?? throw new ArgumentNullException(nameof(ctl));
			ctl.RunWhenHandleReady(tb => WinAPI.SendMessage(
				tb.Handle,
				WinAPI.WindowMessages.EM_SETCUEBANNER,
				0,
				BannerText));
		}

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static void RunWhenHandleReady<T>(this T ctl, Action<Control> HandleReadyAction) where T : Control
		{
			_ = ctl ?? throw new ArgumentNullException(nameof(ctl));
			if (ctl.Disposing || ctl.IsDisposed) return;

			if (ctl.IsHandleCreated)
			{
				HandleReadyAction?.Invoke(ctl);//Control handle already Exist, run immediate
			}
			else
			{
				//Delay action when handle will be ready...
				ctl.HandleCreated += (s, e) => HandleReadyAction?.Invoke((T)s);
			}
		}

		/// <summary>
		/// Usually used when you need to do an action with a slight delay after exiting the current method. 
		/// For example, if some data will be ready only after exiting the control event handler processing branch
		/// </summary>
		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static void RunDelayed(this Action DelayedAction, int DelayInterval = 100)
		{
			_ = DelayedAction ?? throw new ArgumentNullException(nameof(DelayedAction));

			//Use 'System.Windows.Forms.Timer' that uses some thread with caller to raise events
			System.Windows.Forms.Timer tmrDelay = new()
			{
				Interval = DelayInterval,
				Enabled = false //do not start timer untill we finish it's setup
			};
			tmrDelay.Tick += (s, te) =>
			{
				//first stop and dispose our timer, to avoid double execution
				tmrDelay.Stop();
				tmrDelay.Dispose();

				//Now start action
				DelayedAction.Invoke();
			};

			//Start delay timer
			tmrDelay.Start();
		}

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

		/*
		 
		internal static GraphicsPath CreateRoundRect(this Graphics g, Rectangle rect)
		{
			GraphicsPath path = new();
			path.AddArc(rect.Left, rect.Top, rect.Height, rect.Height, 90, 180);
			path.AddArc(rect.Right - rect.Height, rect.Top, rect.Height, rect.Height, -90, 180);
			path.CloseFigure();
			return path;
		}
		*/

		[DebuggerNonUserCode, DebuggerStepThrough, MethodImpl(MethodImplOptions.AggressiveInlining)]
		internal static void DrawRoundRect(
			   this Graphics g,
			   Rectangle rc,
			   int radius,
			   Pen? pen = null,
			   Brush? brush = null)
		{
			Rectangle rcCorner = new(rc.X, rc.Y, radius, radius);

			using (System.Drawing.Drawing2D.GraphicsPath path = new())
			{
				path.AddArc(rcCorner, 180, 90);

				rcCorner.X = rc.X + rc.Width - radius;
				path.AddArc(rcCorner, 270, 90);

				rcCorner.Y = rc.Y + rc.Height - radius;
				path.AddArc(rcCorner, 0, 90);

				rcCorner.X = rc.X;
				path.AddArc(rcCorner, 90, 90);

				path.CloseFigure();

				if (brush != null) g.FillPath(brush, path);
				if (pen != null) g.DrawPath(pen, path);
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


		#region AttachDelayedFilter

		private const int DEFAULT_TEXT_EDIT_DELAY = 1000;
		private const string DEFAULT_FILTER_CUEBANNER = "Filter";

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static void AttachDelayedFilter(
			this TextBox txtCtl,
			Action<string> OnTextChangedCallBack,
			int TextEditiDelay = DEFAULT_TEXT_EDIT_DELAY,
			string VistaCueBanner = DEFAULT_FILTER_CUEBANNER,
			bool SetBackColorAsSystemTipColor = true)
		{
			var TMR = new System.Windows.Forms.Timer() { Interval = TextEditiDelay };
			txtCtl.Tag = TMR; //Сохраняем ссылку на таймер хоть где-то, чтобы GC его не грохнул.

			if (!string.IsNullOrWhiteSpace(VistaCueBanner)) txtCtl.SetVistaCueBanner(VistaCueBanner);
			if (SetBackColorAsSystemTipColor) txtCtl.BackColor = SystemColors.Info;

			TMR.Tick += (s, e) =>
			{
				TMR.Stop(); //Останавливаем таймер
				var sNewText = txtCtl.Text;
				OnTextChangedCallBack.Invoke(sNewText);
			};
			txtCtl.TextChanged += (s, e) =>
			{
				//Перезапускаем таймер
				TMR.Stop();
				TMR.Start();
			};
		}


		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static void AttachDelayedFilter(
			this TextBox txtCtl,
			Action TextChangedCallBack,
			int iDelay_ms = DEFAULT_TEXT_EDIT_DELAY,
			string VistaCueBanner = DEFAULT_FILTER_CUEBANNER,
			bool SetBackColorAsSystemTipColor = true)
		{
			Action<string> DummyCallback = new((s) => TextChangedCallBack.Invoke());
			txtCtl.AttachDelayedFilter(
				DummyCallback,
				iDelay_ms,
				VistaCueBanner,
				SetBackColorAsSystemTipColor);
		}


		/*
		 
	<MethodImpl(MethodImplOptions.AggressiveInlining), System.Runtime.CompilerServices.Extension() >
	Friend Sub AttachDelayedFilter(TB As System.Windows.Forms.ToolStripTextBox,
								   TextChangedCallBack As Action,
								   Optional iDelay_ms As Integer = 1000,
								   Optional VistaCueBanner As String = C_DEFAULT_FILTER_TEXTBOX_CUE_BANNER,
								   Optional SetBackColorAsSystemTipColor As Boolean = True)


		With TB

			Call.TextBox.AttachDelayedFilter(TextChangedCallBack, iDelay_ms, VistaCueBanner, SetBackColorAsSystemTipColor)

		   If(SetBackColorAsSystemTipColor) Then.BackColor = SystemColors.Info
	 End With
 End Sub
		*/



		#endregion

	}
}
