using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using Microsoft.Win32.SafeHandles;

#region Structures to interoperate with the Windows API

using word = System.UInt16;
using dword = System.UInt32;
using hwnd = System.IntPtr;
using large_int = System.Int64;
using ulong_ptr = System.IntPtr;

#endregion

#nullable enable

namespace Navferty.Common.WinAPI.GDI
{
	internal class DC : SafeHandleZeroOrMinusOneIsInvalid
	{
		[DllImport(Core.WINDLL_USER)]
		private static extern IntPtr GetDC(hwnd hwnd);

		[DllImport(Core.WINDLL_USER)]
		private static extern bool ReleaseDC(hwnd hwnd, IntPtr hdc);

		internal hwnd hWnd = IntPtr.Zero;

		public DC(hwnd WindowHandle) : base(true)
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
			return bResult;
		}

		public Graphics CreateGraphics() => Graphics.FromHdc(DangerousGetHandle());
	}

	internal static class WinAPI_Extensions
	{
		public static DC GetDC(this IWin32Window wnd) => new(wnd.Handle);
	}
}
