using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Win32.SafeHandles;

namespace NavfertyExcelAddIn.Commons
{
	internal static class WinAPI
	{
		public enum WindowMessages : int
		{
			WM_PAINT = 0xF,
			EM_SETCUEBANNER = 0x1501
		}

		internal const string WINDLL_USER = "user32.dll";

		[DllImport(WINDLL_USER, SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.Winapi)]
		internal static extern IntPtr SendMessage(
			[In] IntPtr hwnd,
			[In, MarshalAs(UnmanagedType.I4)] WinAPI.WindowMessages wMsg,
			[In] int wParam,
			[In, MarshalAs(UnmanagedType.LPTStr)] string lParam);

		[DllImport(WINDLL_USER)]
		private static extern IntPtr GetDC(
			[In] IntPtr hwnd);

		[DllImport(WINDLL_USER)]
		private static extern int ReleaseDC(
			[In] IntPtr hwnd,
			[In] IntPtr hdc);

		[DllImport(WINDLL_USER)]
		private static extern int GetClientRect(
			[In] IntPtr hwnd,
			[In, Out] ref System.Drawing.Rectangle rc);

		public static Rectangle GetClientRect(IWin32Window wind)
		{
			var rcClient = new Rectangle();
			GetClientRect(wind.Handle, ref rcClient);
			return rcClient;
		}


		internal class DC : SafeHandleZeroOrMinusOneIsInvalid
		{
			[DllImport(WINDLL_USER)]
			private static extern IntPtr GetDC(IntPtr hwnd);

			[DllImport(WINDLL_USER)]
			private static extern bool ReleaseDC(IntPtr hwnd, IntPtr hdc);

			internal IntPtr hWnd = IntPtr.Zero;

			public DC(IntPtr WindowHandle) : base(true)
			{
				var hdc = GetDC(WindowHandle);
				if (hdc == IntPtr.Zero) throw new Win32Exception();
				hWnd = WindowHandle;
				SetHandle(hdc);
			}

			public DC(IWin32Window Window) : this(Window.Handle) { }

			protected override bool ReleaseHandle()
			{
				if (IsInvalid) return true;
				bool bResult = ReleaseDC(hWnd, handle);
				SetHandle(IntPtr.Zero);

				Debug.WriteLine("ReleaseHandle");
				return bResult;
			}

			public Graphics CreateGraphics() => Graphics.FromHdc(DangerousGetHandle());
		}
	}
}
