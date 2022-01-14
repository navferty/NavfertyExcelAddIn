using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
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

		private const string WINDLL_USER = "user32.dll";

		[DllImport(WINDLL_USER, SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.Winapi)]
		private static extern IntPtr SendMessage(
			[In] IntPtr hwnd,
			[In] int wMsg,
			[In] int wParam,
			[In, MarshalAs(UnmanagedType.LPTStr)] string lParam);

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static void SetVistaCueBanner(this TextBox ctl, string BannerText = null)
		{
			_ = ctl ?? throw new ArgumentNullException(nameof(ctl));

			const int EM_SETCUEBANNER = 0x1501;
			//Action<TextBox, string> cbSetBannerWhenHandleReady = (t, s) =>				 SendMessage(t.Handle, EM_SETCUEBANNER, 0, s);
			ctl.RunWhenHandleReady(tb => SendMessage(tb.Handle, EM_SETCUEBANNER, 0, BannerText));
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


	}
}
