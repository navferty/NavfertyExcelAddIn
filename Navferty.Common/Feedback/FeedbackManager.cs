using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Microsoft.Win32;

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

					var sBitmapFile = System.IO.Path.Combine(
						  System.IO.Path.GetTempPath(), (Guid.NewGuid().ToString() + '.'.ToString() + fileExt));
					bmCapt.Save(sBitmapFile, fmt);
					return new System.IO.FileInfo(sBitmapFile);
				}
			}).ToArray();

		private const string MAIL_SUBJECT = @"NavfertyExcelAddin Bug report from user!";
		private const int MAX_MESSAGE_BODY_LENGH = 500;

		private static bool SendFeedEMail(
			string messageBody,
			bool sendScreenshots = true
			)
		{
			//TODO: !!! Insert developer email instead of this !!!
			string developerMail = (new Func<string>(() =>
			{
				using (var hKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676\00000005"))
					return hKey.GetValue("Email").ToString();
			})).Invoke();


			List<FileInfo> lScreenshotFiles = new();

			if (string.IsNullOrWhiteSpace(messageBody))
				messageBody = string.Empty;



			if (messageBody.Length > MAX_MESSAGE_BODY_LENGH)
				messageBody = new string(messageBody.Take(MAX_MESSAGE_BODY_LENGH).ToArray());


			StringBuilder sbTechInfo = new();
			sbTechInfo.AppendLine($"Message created {DateTime.Now}");
			sbTechInfo.AppendLine(Application.ProductVersion.ToString());
			sbTechInfo.AppendLine(Application.ExecutablePath);
			sbTechInfo.AppendLine();
			sbTechInfo.AppendLine(messageBody);

			messageBody = sbTechInfo.ToString();

			if (sendScreenshots)    //Create screenshots to temp dir
				lScreenshotFiles = GetScreenshotsAsFiles(ImageFormat.Jpeg).ToList();

			try
			{
				//Send Screenshots 
				var bSend = WinAPI.MAPI.SendMail(
					developerMail,
					MAIL_SUBJECT,
					messageBody,
					WinAPI.MAPI.UIFlags.SendMailDirectNoUI,
					lScreenshotFiles.Select(fi => fi.FullName).ToArray()
					);

				return bSend;
			}
			finally
			{
				//Cleanup Temp files
				lScreenshotFiles.ForEach(fi =>
				{
					try { fi.Delete(); }
					catch { }
				});
			}
		}

		public static void ShowFeedbackUI()
		{
			SendFeedEMail("sd", true);

		}
	}
}
