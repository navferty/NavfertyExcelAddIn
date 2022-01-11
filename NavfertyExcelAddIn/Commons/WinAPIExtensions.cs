using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using Microsoft.Office.Core;

using NavfertyExcelAddIn.Localization;

using Application = Microsoft.Office.Interop.Excel.Application;

namespace NavfertyExcelAddIn.Commons
{
	[DebuggerStepThrough]
	internal static class WinAPIExtensions
	{

		private const string WINDLL_USER = "user32.dll";

		[DllImport(WINDLL_USER, SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.Winapi)]
		private static extern IntPtr SendMessage(
			[In] IntPtr hwnd,
			[In] int wMsg,
			[In] int wParam,
			[In, MarshalAs(UnmanagedType.LPTStr)] string lParam);


		public static void SetVistaCueBanner(this TextBox ctl, string BannerText = null)
		{
			_ = ctl ?? throw new ArgumentNullException(nameof(ctl));

			const int EM_SETCUEBANNER = 0x1501;
			Action<TextBox, string> cbSetBannerWhenHandleReady = (t, s) =>
				 SendMessage(t.Handle, EM_SETCUEBANNER, 0, s);

			if (ctl.IsHandleCreated)
			{
				cbSetBannerWhenHandleReady(ctl, BannerText);
			}
			else
			{
				ctl.HandleCreated += (s, e) => cbSetBannerWhenHandleReady(s as TextBox, BannerText);
			}
		}

		public static void SetVistaCueBanner(this ComboBox ctl, string BannerText = null)
		{
			_ = ctl ?? throw new ArgumentNullException(nameof(ctl));

			const int CB_SETCUEBANNER = 0x1703;

			Action<ComboBox, string> cbSetBannerWhenHandleReady = (t, s) =>
					 SendMessage(t.Handle, CB_SETCUEBANNER, 0, s);

			if (ctl.IsHandleCreated)
			{
				cbSetBannerWhenHandleReady(ctl, BannerText);
			}
			else
			{
				ctl.HandleCreated += (s, e) => cbSetBannerWhenHandleReady(s as ComboBox, BannerText);
			}
		}


	}
}
