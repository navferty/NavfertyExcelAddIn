using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;

using Microsoft.Win32;

using NLog;

#nullable enable

namespace Navferty.Common.Feedback
{
	public static class FeedbackManager
	{
		private const string GITHUB_BUGTRACKER_URL = @"https://github.com/navferty/NavfertyExcelAddIn/issues";
		private const string DEVELOPER_MAIL = @"navferty@ymail.com";
		private const string MAIL_SUBJECT = @"NavfertyExcelAddin Bug report from user!";
		internal const int MAX_USER_TEXT_LENGH = 1_000;

		private static FileInfo[] GetScreenshotsAsFiles(ImageFormat fmt, string fileExt = "jpg")
			=> System.Windows.Forms.Screen.AllScreens.ToList().Select(scr =>
			{
				using (Bitmap bmCapt = new(scr.Bounds.Width, scr.Bounds.Height, PixelFormat.Format32bppArgb))
				{
					var rcCapt = scr.Bounds;
					using (Graphics g = Graphics.FromImage(bmCapt))
						g.CopyFromScreen(rcCapt.Left, rcCapt.Top, 0, 0, rcCapt.Size);

					var sBitmapFile = Path.Combine(Path.GetTempPath(), (Guid.NewGuid().ToString() + '.'.ToString() + fileExt));
					bmCapt.Save(sBitmapFile, fmt);
					return new System.IO.FileInfo(sBitmapFile);
				}
			}).ToArray();

		/// <summary>This used for debug!
		//If you want to send error reports to custom email,
		//create new string value 'Navferty_ExcelAddIn_Feedback_Email' in root of 'HKEY_CURRENT_USER' and set you mail
		/// </summary>
		internal static string GetDeveloperMail()
		{
			try
			{
				string mail = Registry.CurrentUser.GetValue("Navferty_ExcelAddIn_Feedback_Email").ToString().Trim();
				if (string.IsNullOrWhiteSpace(mail)) mail = DEVELOPER_MAIL;
				return mail;
			}
			catch { return DEVELOPER_MAIL; }
		}

		internal static bool SendFeedEMail(
			string userText,
			bool sendScreenshots = true,
			Form? parentWindow = null
			)
		{

			var logger = LogManager.GetCurrentClassLogger();
			logger.Debug("Start SendFeedEMail Task...");

			string developerMail = GetDeveloperMail();
			logger.Debug($"developerMail: '{developerMail}'");

			List<FileInfo> lFilesToAttach = new();
			if (userText.Length > MAX_USER_TEXT_LENGH) userText = new string(userText.Take(MAX_USER_TEXT_LENGH).ToArray());

			string sysInfo = GetSystemInfo().Trim();
			logger.Debug($"System Info Dump:\n{sysInfo}\n\n********\nUser message: '{userText}'\n");

			StringBuilder sbMessageBody = new();
			sbMessageBody.Append("User message: ");
			sbMessageBody.AppendLine((string.IsNullOrWhiteSpace(userText) ? "[NONE]" : ('"' + userText + '"')));
			string messageBody = sbMessageBody.ToString();

			if (sendScreenshots)
			{
				if (parentWindow != null)
				{
					//Temporary hide feedback UI to make clear screenshots
					parentWindow.Opacity = 0;
					parentWindow.Refresh();
					Application.DoEvents();
				}
				try
				{
					//Create screenshots to temp dir
					var screenshotFiles = GetScreenshotsAsFiles(ImageFormat.Jpeg);
					lFilesToAttach.AddRange(screenshotFiles);
					string screenshotFileNames = string.Join(", ", screenshotFiles.Select(fi => fi.FullName));
					logger.Debug($"Screenshot Files ({lFilesToAttach.Count}): '{screenshotFileNames}'");
				}
				finally
				{
					if (parentWindow != null)
					{
						//Restore feedback UI
						parentWindow.Opacity = 1;
						parentWindow.Refresh();
						Application.DoEvents();
					}
				}
			}

			var fiNLogFile = GetNLogFile();
			if (null != fiNLogFile && fiNLogFile.Exists)
			{
				logger.Debug($"Log File Found: '{fiNLogFile.FullName}', Exist: {fiNLogFile.Exists}");
				FileInfo fiLogFileInTempDir = new(Path.Combine(
					Path.GetTempPath(),
					(Guid.NewGuid().ToString() + "_" + fiNLogFile.Name)));

				//Drop last NLog cache data to disk
				{
					//logger.Factory.Flush();
					LogManager.Flush();
					Thread.Sleep(1000); //Waiting NLog flush task to finish
				}

				//Copy NLog file to temp file
				fiNLogFile.CopyTo(fiLogFileInTempDir.FullName);

				//Attach temp NLog file to email
				if (fiLogFileInTempDir.Exists) lFilesToAttach.Add(fiLogFileInTempDir);
			}

			logger.Debug($"Total Files To Attach: '{lFilesToAttach.Count}'");
			try
			{
				//Send Screenshots 
				var bSend = WinAPI.MAPI.SendMail(
					developerMail,
					MAIL_SUBJECT,
					messageBody,
					WinAPI.MAPI.UIFlags.SendMailDirectNoUI,
					parentWindow,
					lFilesToAttach.Select(fi => fi.FullName).ToArray()
					);

				logger.Debug($"Send result: {bSend}");
				return bSend;
			}
			finally
			{
				//Cleanup Temp files
				lFilesToAttach.ForEach(fi =>
				{
					try { fi.Delete(); }
					catch { }
				});
			}
		}
		private static string GetSystemInfo()
		{
			Func<Assembly, string?, string> DumpAssemmbly = new((asm, title) =>
			{
				StringBuilder sbAsm = new();
				sbAsm.AppendLine($"*** {title ?? string.Empty} Assembly '{asm.FullName}'");
				sbAsm.AppendLine($"Location: '{asm.Location}'");
				sbAsm.AppendLine($"ImageRuntimeVersion: '{asm.ImageRuntimeVersion}'");
				sbAsm.AppendLine($"Trusted: '{asm.IsFullyTrusted}'");
				sbAsm.AppendLine($"EntryPoint: '{asm.EntryPoint}'");
				return sbAsm.ToString();
			});

			Func<AssemblyName, string?, string> DumpAssemmblyName = new((asmn, title) =>
			{
				StringBuilder sbAsm = new();
				sbAsm.AppendLine($"*** {title ?? string.Empty} Assembly '{asmn.FullName}'");
				sbAsm.AppendLine($"CodeBase: '{asmn.CodeBase}'");
				sbAsm.AppendLine($"ContentType: '{asmn.ContentType}'");
				sbAsm.AppendLine($"Culture: '{asmn.CultureInfo.DisplayName}'");
				sbAsm.AppendLine($"ProcessorArchitecture: '{asmn.ProcessorArchitecture}'");
				return sbAsm.ToString();
			});

			var dtNow = DateTime.Now;
			var asm = Assembly.GetExecutingAssembly();
			StringBuilder sbSysInfo = new();
			sbSysInfo.AppendLine("*** Product:");
			sbSysInfo.AppendLine($"Name: '{Application.ProductName}' v'{Application.ProductVersion}'");
			sbSysInfo.AppendLine($"Path: '{Application.ExecutablePath}'");
			sbSysInfo.AppendLine();
			sbSysInfo.AppendLine(DumpAssemmbly(Assembly.GetExecutingAssembly(), "Executing"));
			sbSysInfo.AppendLine(DumpAssemmbly(Assembly.GetCallingAssembly(), "Calling"));

			Assembly.GetExecutingAssembly()
				.GetReferencedAssemblies()
				.OrderBy(asmn => asmn.FullName)
				.ToList()
				.ForEach(asmn => { sbSysInfo.AppendLine(DumpAssemmblyName(asmn, "Referenced")); });

			sbSysInfo.AppendLine("*** TimeZone:");
			sbSysInfo.AppendLine(dtNow.Kind.ToString() + $": {dtNow}");
			sbSysInfo.AppendLine($"Utc: {dtNow.ToUniversalTime()}");
			sbSysInfo.AppendLine($"UtcOffset: {TimeZone.CurrentTimeZone.GetUtcOffset(dtNow)}");
			sbSysInfo.AppendLine();
			sbSysInfo.AppendLine("*** Culture:");
			sbSysInfo.AppendLine($"CultureInfo.CurrentCulture: {CultureInfo.CurrentCulture}");
			sbSysInfo.AppendLine($"CultureInfo.CurrentUICulture: {CultureInfo.CurrentUICulture}");
			sbSysInfo.AppendLine($"Application.CurrentCulture: {Application.CurrentCulture}");
			sbSysInfo.AppendLine($"InputLanguage: {Application.CurrentInputLanguage.Culture} (Layout: {Application.CurrentInputLanguage.LayoutName})");
			sbSysInfo.AppendLine();
			sbSysInfo.AppendLine($"VisualStyleState: {Application.VisualStyleState}");
			return sbSysInfo.ToString();
		}

		private static FileInfo? GetNLogFile()
		{
			LogManager.Flush(); //Write NLog cache to disk if this still in RAM
			var logFileName = LogManagement.GetTargetFilename(null);
			if (logFileName != null) return new(logFileName);
			return null;
		}

		public static void ShowFeedbackUI()
		{
			using var fui = new frmFeedbackUI();
			fui.ShowDialog();
		}

		public static void ShowGithub()
			=> System.Diagnostics.Process.Start(GITHUB_BUGTRACKER_URL);

		public static void ShowLogFile()
			=> System.Diagnostics.Process.Start(GetNLogFile()!.FullName);

	}
}
