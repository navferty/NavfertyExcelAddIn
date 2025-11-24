using System;
using System.Runtime.InteropServices;
using System.Threading;

using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using Application = Microsoft.Office.Interop.Excel.Application;

namespace NavfertyExcelAddIn.UnitTests.SqliteExport;

public class AutomationTestsBase : TestsBase
{
	protected readonly TimeSpan defaultSleep = TimeSpan.FromMilliseconds(750);
	protected Application app;

	[TestInitialize]
	public virtual void Initialize()
	{
		app = OpenNewExcelApp();
		app.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized;

		// Force Excel to become the foreground window
		var excelWindowHandle = new IntPtr(app.Hwnd);
		BringWindowToForeground(excelWindowHandle);
		Thread.Sleep(defaultSleep);

		// check current active window to be excel app
		var currentActiveWindowHandle = GetForegroundWindow();

		TestContext.WriteLine($"Excel window handle: {excelWindowHandle}");
		TestContext.WriteLine($"Current active window handle: {currentActiveWindowHandle}");
		Assert.AreEqual(excelWindowHandle, currentActiveWindowHandle, "Excel is not the active window");
	}

	[TestCleanup]
	public virtual void Cleanup()
	{
		app.Quit();
	}

	private static Application OpenNewExcelApp()
	{
		return new Application
		{
			Visible = true,
			EnableEvents = false,
			DisplayAlerts = false,
			AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable
		};
	}

	protected static void BringWindowToForeground(IntPtr windowHandle)
	{
		ShowWindow(windowHandle, SW_RESTORE);
		SetForegroundWindow(windowHandle);
	}

	[DllImport("user32.dll")]
	protected static extern IntPtr GetForegroundWindow();

	[DllImport("user32.dll")]
	[return: MarshalAs(UnmanagedType.Bool)]
	private static extern bool SetForegroundWindow(IntPtr hWnd);

	[DllImport("user32.dll")]
	[return: MarshalAs(UnmanagedType.Bool)]
	private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

	[DllImport("user32.dll")]
	[return: MarshalAs(UnmanagedType.Bool)]
	private static extern bool BringWindowToTop(IntPtr hWnd);

	[DllImport("user32.dll")]
	private static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr processId);

	[DllImport("kernel32.dll")]
	private static extern uint GetCurrentThreadId();

	[DllImport("user32.dll")]
	[return: MarshalAs(UnmanagedType.Bool)]
	private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

	[DllImport("user32.dll", SetLastError = true)]
	private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

	[DllImport("user32.dll", SetLastError = true)]
	private static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);

	[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
	private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

	[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
	private static extern int GetWindowTextLength(IntPtr hWnd);

	private const int SW_RESTORE = 9;
	private const int SW_SHOW = 5;

	protected string GetMessageBoxText()
	{
		var timeoutMs = 5000;
		var startTime = DateTime.Now;

		while ((DateTime.Now - startTime).TotalMilliseconds < timeoutMs)
		{
			Thread.Sleep(100);

			// Try to find message box with title "Info"
			var msgBoxHandle = FindWindow("#32770", "Info");
			if (msgBoxHandle == IntPtr.Zero)
				continue;

			// Skip the first Static control (usually the icon) and get the second one (message text)
			var staticHandle = FindWindowEx(msgBoxHandle, IntPtr.Zero, "Static", null);
			staticHandle = FindWindowEx(msgBoxHandle, staticHandle, "Static", null);

			if (staticHandle == IntPtr.Zero)
				continue;

			var length = GetWindowTextLength(staticHandle);
			if (length == 0)
				continue;

			var sb = new System.Text.StringBuilder(length + 1);
			GetWindowText(staticHandle, sb, sb.Capacity);
			return sb.ToString();
		}

		return string.Empty;
	}
}
