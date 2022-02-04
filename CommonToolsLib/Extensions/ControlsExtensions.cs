using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;


namespace NavfertyCommon
{
	[DebuggerStepThrough]
	public static class ControlsExtensions
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

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static StringFormat ToStringFormat(this ContentAlignment ca)
		{
			var sf = new StringFormat()
			{
				Alignment = ca.GetAlignment(),
				LineAlignment = ca.GetLineAlignment()
			};
			return sf;
		}

		#region AttachDelayedFilter

		private const int DEFAULT_TEXT_EDIT_DELAY = 1000;
		private const string DEFAULT_FILTER_CUEBANNER = "Filter";

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static void AttachDelayedFilter(
			this TextBox txtCtl,
			Action<string> OnTextChangedCallBack,
			int TextEditiDelay = DEFAULT_TEXT_EDIT_DELAY,
			string VistaCueBanner = DEFAULT_FILTER_CUEBANNER,
			bool SetBackColorAsSystemTipColor = true)
		{
			var TMR = new System.Windows.Forms.Timer() { Interval = TextEditiDelay };
			txtCtl.Tag = TMR; //Сохраняем ссылку на таймер хоть где-то, чтобы GC его не грохнул.

			if (!string.IsNullOrWhiteSpace(VistaCueBanner)) txtCtl.SetVistaCueBanner(VistaCueBanner);
			if (SetBackColorAsSystemTipColor) txtCtl.BackColor = SystemColors.Info;

			TMR.Tick += (s, e) =>
			{
				TMR.Stop(); //Останавливаем таймер
				var sNewText = txtCtl.Text;
				OnTextChangedCallBack.Invoke(sNewText);
			};
			txtCtl.TextChanged += (s, e) =>
			{
				//Перезапускаем таймер
				TMR.Stop();
				TMR.Start();
			};
		}


		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static void AttachDelayedFilter(
			this TextBox txtCtl,
			Action TextChangedCallBack,
			int iDelay_ms = DEFAULT_TEXT_EDIT_DELAY,
			string VistaCueBanner = DEFAULT_FILTER_CUEBANNER,
			bool SetBackColorAsSystemTipColor = true)
		{
			Action<string> DummyCallback = new((s) => TextChangedCallBack.Invoke());
			txtCtl.AttachDelayedFilter(
				DummyCallback,
				iDelay_ms,
				VistaCueBanner,
				SetBackColorAsSystemTipColor);
		}


		/*
		 
	<MethodImpl(MethodImplOptions.AggressiveInlining), System.Runtime.CompilerServices.Extension() >
	Friend Sub AttachDelayedFilter(TB As System.Windows.Forms.ToolStripTextBox,
								   TextChangedCallBack As Action,
								   Optional iDelay_ms As Integer = 1000,
								   Optional VistaCueBanner As String = C_DEFAULT_FILTER_TEXTBOX_CUE_BANNER,
								   Optional SetBackColorAsSystemTipColor As Boolean = True)


		With TB

			Call.TextBox.AttachDelayedFilter(TextChangedCallBack, iDelay_ms, VistaCueBanner, SetBackColorAsSystemTipColor)

		   If(SetBackColorAsSystemTipColor) Then.BackColor = SystemColors.Info
	 End With
 End Sub
		*/



		#endregion

	}
}
