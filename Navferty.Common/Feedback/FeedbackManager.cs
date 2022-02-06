using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

using Microsoft.Win32;

using NLog;

#nullable enable

namespace Navferty.Common.Feedback
{
	public static class FeedbackManager
	{
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

		private const string MAIL_SUBJECT = @"NavfertyExcelAddin Bug report from user!";
		private const int MAX_MESSAGE_BODY_LENGH = 1024 * 3;

		private static readonly Lazy<ILogger> logger = new(() => LogManager.GetCurrentClassLogger());

		internal static bool SendFeedEMail(
			string userText,
			bool sendScreenshots = true
			)
		{
			List<FileInfo> lFilesToAttach = new();

			logger.Value.Debug("Start SendFeedEMail");
			var logFileName = LogManagement.GetTargetFilename(null);
			var fiLogOld = new FileInfo(logFileName);
			if (fiLogOld.Exists)
			{
				logger.Value.Debug($"Attaching Log File: '{fiLogOld.FullName}'");
				var sNewLogFile = Path.Combine(Path.GetTempPath(), (Guid.NewGuid().ToString() + "_" + fiLogOld.Name));
				var fiLogNew = new FileInfo(sNewLogFile);
				fiLogOld.CopyTo(fiLogNew.FullName);
				if (fiLogNew.Exists) lFilesToAttach.Add(fiLogNew);
			}

			//TODO: !!! Insert developer email instead of this !!!
			string developerMail = (new Func<string>(() =>
			{
				using (var hKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676\00000005"))
					return hKey.GetValue("Email").ToString();
			})).Invoke();

			logger.Value.Debug($"developerMail: '{developerMail}'");

			StringBuilder sbMessageBody = new();
			sbMessageBody.AppendLine(GetSystemInfo().Trim());
			sbMessageBody.AppendLine("*** User message:");
			sbMessageBody.AppendLine((string.IsNullOrWhiteSpace(userText) ? "[NONE]" : ('"' + userText + '"')));

			string messageBody = sbMessageBody.ToString();
			if (messageBody.Length > MAX_MESSAGE_BODY_LENGH) messageBody = new string(messageBody.Take(MAX_MESSAGE_BODY_LENGH).ToArray());
			logger.Value.Debug($"messageBody:\n'{messageBody}'");

			if (sendScreenshots) lFilesToAttach = GetScreenshotsAsFiles(ImageFormat.Jpeg).ToList();//Create screenshots to temp dir
			logger.Value.Debug($"FilesToAttach: '{lFilesToAttach.Count}'");


			try
			{
				//Send Screenshots 
				var bSend = WinAPI.MAPI.SendMail(
					developerMail,
					MAIL_SUBJECT,
					messageBody,
					WinAPI.MAPI.UIFlags.SendMailDirectNoUI,
					lFilesToAttach.Select(fi => fi.FullName).ToArray()
					);

				logger.Value.Debug($"Send result: {bSend}");
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
			var dtNow = DateTime.Now;
			var asm = Assembly.GetExecutingAssembly();
			StringBuilder sbSysInfo = new();
			sbSysInfo.AppendLine("*** Product:");
			sbSysInfo.AppendLine($"Name: '{Application.ProductName}' v'{Application.ProductVersion}'");
			sbSysInfo.AppendLine($"Path: '{Application.ExecutablePath}'");
			sbSysInfo.AppendLine();
			sbSysInfo.AppendLine("*** Assembly:");
			sbSysInfo.AppendLine($"FullName: '{asm.FullName}'");
			sbSysInfo.AppendLine($"Location: '{asm.Location}'");
			sbSysInfo.AppendLine($"ImageRuntimeVersion: '{asm.ImageRuntimeVersion}'");
			sbSysInfo.AppendLine($"IsFullyTrusted: '{asm.IsFullyTrusted}'");
			sbSysInfo.AppendLine($"EntryPoint: '{asm.EntryPoint}'");
			sbSysInfo.AppendLine();
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

		public static void ShowFeedbackUI()
		{
			using var fui = new frmFeedbackUI();
			fui.ShowDialog();
		}
	}
}
