using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NavfertyExcelAddIn.Commons
{
	internal static class ControlsExtensions_2
	{


		#region AttachDelayedFilter

		private const int DEFAULT_TEXT_EDIT_DELAY = 1000;
		private const string DEFAULT_FILTER_CUEBANNER = "Фильтр";

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

			//If(!string.IsNullOrWhiteSpace(VistaCueBanner)) Then Call.ExtCtl_Vista_SetCueBanner(VistaCueBanner)
			//If(SetBackColorAsSystemTipColor) Then.BackColor = SystemColors.Info

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
