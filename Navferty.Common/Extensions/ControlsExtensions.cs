using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

using Navferty.Common.WinAPI;

#nullable enable

namespace Navferty.Common
{
	[DebuggerStepThrough]
	public static class ControlsExtensions
	{
		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static void SetVistaCueBanner(this TextBox ctl, string? BannerText = null)
		{
			_ = ctl ?? throw new ArgumentNullException(nameof(ctl));
			ctl.RunWhenHandleReady(tb => Windows.SendMessage(
				tb.Handle,
				Windows.WindowMessages.EM_SETCUEBANNER,
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
				ctl.HandleCreated += (s, e) => HandleReadyAction?.Invoke((T)s);
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


		#region AttachDelayedFilter

		private const int DEFAULT_TEXT_EDIT_DELAY = 1000;
		private const string DEFAULT_FILTER_CUEBANNER = "Filter";

		/// <summary>
		/// Attaches a deferred text change event handler that makes it possible to react to text changes with some delay, 
		/// allowing the user to correct erroneous input, 
		/// or complete input, rather than reacting immediately to each letter.
		/// </summary>
		/// <param name="OnTextChangedCallBack">TextChanged Handler</param>
		/// <param name="TextEditiDelay">Delay (ms.) during which repeated input will not call the handler</param>
		/// <param name="VistaCueBanner">Vista cueabanner text</param>
		/// <param name="SetBackColorAsSystemTipColor">Sets the background color for textbox to Systemcolors.Info</param>
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
				//Restart timer...
				TMR.Stop();
				TMR.Start();
			};
		}

		/// <inheritdoc cref="AttachDelayedFilter" />
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

		#endregion

	}
}
