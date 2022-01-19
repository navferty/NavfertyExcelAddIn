using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using Microsoft.Office.Core;

using NavfertyExcelAddIn.Localization;

using Application = Microsoft.Office.Interop.Excel.Application;

namespace NavfertyExcelAddIn.Commons
{
	[DebuggerStepThrough]
	internal static class ControlsExtensions
	{
		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static void SetVistaCueBanner(this TextBox ctl, string BannerText = null)
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
				ctl.HandleCreated += (s, e) => HandleReadyAction?.Invoke(s as T);
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
		public static StringAlignment GetAlignment(this ContentAlignment ca)
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
		public static StringAlignment GetLineAlignment(this ContentAlignment ca)
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
	}
}
