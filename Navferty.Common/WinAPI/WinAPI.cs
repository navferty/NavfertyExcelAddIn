using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using Microsoft.Win32.SafeHandles;

#nullable enable

namespace Navferty.Common
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
			[In, MarshalAs(UnmanagedType.LPTStr)] string? lParam);

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
				return bResult;
			}

			public Graphics CreateGraphics() => Graphics.FromHdc(DangerousGetHandle());
		}


		/// <summary>
		/// https://docs.microsoft.com/ru-ru/previous-versions/windows/desktop/windowsmapi/mapi32-dll-stub-registry-settings?redirectedfrom=MSDN
		/// </summary>
		internal static class MAPI
		{
			public const int MAX_ATTACHMENTS = 20;

			public enum MAPI_ERRORS : int
			{
				[Description("OK")]
				OK = 0,

				[Description("User abort")]
				UserAbort = 1,

				[Description("General MAPI failure")]
				GeneralMAPIFailure = 2,

				[Description("MAPI login failure")]
				MAPILoginFailure = 3,

				[Description("Disk full")]
				DiskFull = 4,

				[Description("Insufficient memory")]
				InsufficientMemory = 5,

				[Description("Access denied [6")]
				AccessDenied = 6,

				[Description("Unknown")]
				Unknown = 7,

				[Description("Too many sessions")]
				TooManySessions = 8,

				[Description("Too many files were specified")]
				TooManyFilesWereSpecified = 9,

				[Description("Too many recipients were specified")]
				ToomanyRecipientsWereSpecified = 10,

				[Description("A specified attachment was not found")]
				SpecifiedAttachmentWasNotFound = 11,

				[Description("Attachment open failure")]
				AttachmentOpenFailure = 12,

				[Description("Attachment write failure")]
				AttachmentWriteFailure = 13,

				[Description("Unknown recipient")]
				UnknownRecipient = 14,

				[Description("Bad recipient type")]
				BadRecipientType = 15,

				[Description("No messages")]
				NoMessages = 16,

				[Description("Invalid message")]
				InvalidMessage = 17,

				[Description("Text too large")]
				TextTooLarge = 18,

				[Description("Invalid session")]
				InvalidSession = 19,

				[Description("Type not supported")]
				TypeNotSupported = 20,

				[Description("A recipient was specified ambiguously")]
				RecipientWasSpecifiedAmbiguously = 21,

				[Description("Message in use")]
				MessageInUse = 22,

				[Description("Network failure")]
				NetworkFailure = 23,

				[Description("Invalid edit fields")]
				InvalidEditFields = 24,

				[Description("Invalid recipients")]
				InvalidRecipients = 25,

				[Description("Not supported")]
				NotSupported = 26,

				//[EditorBrowsable(EditorBrowsableState.Never)]
				Last
			}

			public class MAPIException : System.Exception
			{
				public readonly MAPI_ERRORS ErrorCode = MAPI_ERRORS.OK;
				internal MAPIException(MAPI_ERRORS err) : base(GetErrorMessageForCode(err))
				{
					ErrorCode = err;
				}
				public override string Message => GetErrorMessageForCode(ErrorCode);

				public static string GetErrorMessageForCode(MAPI_ERRORS Err)
				{
					if (Err >= MAPI_ERRORS.Last) return $"MAPI Error {Err}";
					return Err.ToString();
				}
			}

			public enum SendToFlags : int
			{
				MAPI_ORIG = 0,
				MAPI_TO,
				MAPI_CC,
				MAPI_BCC
			};

			public enum UIFlags : uint
			{
				/// <summary>Display Send Mail UI dialog</summary>
				PopupUI = 0,

				/// <summary>Send mail without displaying the mail UI dialog.
				/// for safety reason, MSOutlook displays warning UI that allows user to allow or block sending this mail (mail message itself is not visible to user)</summary>
				SendMailDirectNoUI
			}

			#region API

			private enum MAPI_FLAGS : uint
			{
				MAPI_LOGON_UI = 1,              // Показать интерфейс входа в систему
				MAPI_NEW_SESSION = 2,           // Не использовать общий сеанс
				MAPI_DIALOG = 8,
				//MAPI_ALLOW_OTHERS = 8; исходное шестнадцатеричное значение 0x00000008 ; Сделать это общим сеансом
				MAPI_EXPLICIT_PROFILE = 16,     //Не использовать профиль по умолчанию
				MAPI_EXTENDED = 32,             //Расширенный вход MAPI
				MAPI_FORCE_DOWNLOAD = 4096,     //Получить новую почту перед возвратом
				MAPI_SERVICE_UI_ALWAYS = 8192,  //Выполнить вход в систему во всех провайдерах
				MAPI_NO_MAIL = 32768,           //Не активировать транспорт
				MAPI_PASSWORD_UI = 131072,      //Отображать только пользовательский интерфейс пароля
				MAPI_TIMEOUT_SHORT = 1048576,    //Минимальное ожидание ресурсов входа в систему
				MAPI_UNICODE = 0x80000000,
				MAPI_USE_DEFAULT = 0x00000040
			}

			/// <summary>
			/// If you're application is running with Elevated Privileges (i.e. as Administrator) and Outlook isn't, the send will fail. 
			/// You will have to tell the user to close all running instances of Outlook and try again.
			/// </summary>
			[DllImport("MAPI32.DLL", CharSet = CharSet.Ansi)]
			private static extern MAPI_ERRORS MAPISendMail(
				IntPtr lhSession,
				IntPtr hWnd,
				MapiMessage message,
				MAPI_FLAGS flg,
				int rsv);

			#endregion


			public static bool SendMail(
				IEnumerable<MapiRecipDesc> recipients,
				string strSubject,
				string strBody,
				UIFlags UIflags = UIFlags.PopupUI,
				IWin32Window? parentWindow = null,
				params string[] attachFiles
				)
			{
				using (MapiMessage msg = new()
				{
					subject = strSubject,
					noteText = strBody
				})
				{
					(msg.hRecips, msg.recipCount) = RecipientsToBuffer(recipients);
					(msg.hFiles, msg.fileCount) = AttachmentsToBuffer(attachFiles);

					MAPI_FLAGS how = (UIflags == UIFlags.PopupUI) ? (MAPI_FLAGS.MAPI_LOGON_UI | MAPI_FLAGS.MAPI_DIALOG) : (MAPI_FLAGS.MAPI_LOGON_UI);
					how |= MAPI_FLAGS.MAPI_NEW_SESSION;

					IntPtr hWnd = (parentWindow == null) ? IntPtr.Zero : parentWindow.Handle;

					var Err = MAPISendMail(
						new IntPtr(0),
						hWnd,
						msg,
						how,
						0);

					switch (Err)
					{
						case MAPI_ERRORS.OK: return true;
						case MAPI_ERRORS.UserAbort: return false;
						case MAPI_ERRORS.NotSupported: return false;//If send with no UI, and user denied send, this occurs.
						default: throw new MAPIException(Err);
					}
				}
			}

			public static bool SendMail(
				string sendTo,
				string strSubject,
				string strBody,
				UIFlags uiFlags = UIFlags.PopupUI,
				IWin32Window? parentWindow = null,
				params string[] attachFiles)
				=> SendMail(
					new MapiRecipDesc[] { CreateRecipient(sendTo, SendToFlags.MAPI_TO) },
					strSubject, strBody, uiFlags, parentWindow, attachFiles);


			public static MapiRecipDesc CreateRecipient(string email, SendToFlags howTo = SendToFlags.MAPI_TO)
				=> new(email, howTo);

			private static (IntPtr hMem, int Count) RecipientsToBuffer(IEnumerable<MapiRecipDesc> recipients)
			{
				if (!recipients.Any()) throw new ArgumentNullException(nameof(recipients));

				var lRecipients = recipients.ToList();

				int size = Marshal.SizeOf(typeof(MapiRecipDesc));
				IntPtr hMem = Marshal.AllocHGlobal(lRecipients.Count * size);
				IntPtr hWrite = hMem;
				lRecipients.ForEach(mapiDesc =>
				{
					Marshal.StructureToPtr(mapiDesc, hWrite, false);
					hWrite += size;
				});
				return (hMem, lRecipients.Count);
			}

			private static (IntPtr hMem, int fileCount) AttachmentsToBuffer(IEnumerable<string>? attachments = null)
			{
				if (attachments == null || !attachments.Any()) return (IntPtr.Zero, 0);
				if (attachments.Count() > MAX_ATTACHMENTS) throw new ArgumentOutOfRangeException($"{nameof(attachments)}.Count > {MAX_ATTACHMENTS}");

				List<string> lAttachments = attachments.ToList();

				int size = Marshal.SizeOf(typeof(MapiFileDesc));
				IntPtr hMem = Marshal.AllocHGlobal(lAttachments.Count * size);

				MapiFileDesc mapiFileDesc = new();
				mapiFileDesc.position = -1;

				IntPtr ptr = hMem;
				lAttachments.ForEach(path =>
				{
					mapiFileDesc.name = Path.GetFileName(path);
					mapiFileDesc.path = path;
					Marshal.StructureToPtr(mapiFileDesc, ptr, false);
					ptr += size;
				});
				return (hMem, lAttachments.Count);
			}



			[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
			internal class MapiMessage : IDisposable
			{
				public int reserved;
				public string subject;
				public string noteText;
				public string messageType;
				public string dateReceived;
				public string conversationID;
				public int flags;
				public IntPtr hOriginator;
				public int recipCount;
				public IntPtr hRecips;
				public int fileCount;
				public IntPtr hFiles;

				public void Dispose()
				{
					if (hRecips != IntPtr.Zero)
					{
						int size = Marshal.SizeOf(typeof(MapiRecipDesc));
						IntPtr ptr = hRecips;
						for (int i = 0; i < recipCount; i++)
						{
							Marshal.DestroyStructure(ptr, typeof(MapiRecipDesc));
							ptr += size;
						}
						Marshal.FreeHGlobal(hRecips);
					}

					if (hFiles != IntPtr.Zero)
					{
						int size = Marshal.SizeOf(typeof(MapiFileDesc));
						IntPtr ptr = hFiles;
						for (int i = 0; i < fileCount; i++)
						{
							Marshal.DestroyStructure((IntPtr)ptr,
								typeof(MapiFileDesc));
							ptr += size;
						}
						Marshal.FreeHGlobal(hFiles);
					}
				}
			}

			[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
			internal class MapiFileDesc
			{
				public int reserved;
				public int flags;
				public int position;
				public string path;
				public string name;
				public IntPtr hType;
			}

			[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
			public class MapiRecipDesc
			{
				public int reserved = 0;
				public SendToFlags recipClass = SendToFlags.MAPI_TO;

				/// <summary>MAPISendMail hanging with Outlook because the mail addresses had trailing spaces!</summary>
				public string name = String.Empty;
				public string address = String.Empty;
				public int eIDSize = 0;
				public IntPtr hEntryID = IntPtr.Zero;

				private MapiRecipDesc() { }

				public MapiRecipDesc(string email, SendToFlags howTo = SendToFlags.MAPI_TO) : this()
				{
					recipClass = howTo;
					name = email.Trim();
				}
			}
		}
	}

}
