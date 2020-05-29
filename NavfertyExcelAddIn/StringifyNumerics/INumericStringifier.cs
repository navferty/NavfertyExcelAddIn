namespace NavfertyExcelAddIn.StringifyNumerics
{
	public interface INumericStringifier
	{
		/// <summary>
		/// Convert number to text.
		/// </summary>
		/// <param name="number"></param>
		/// <returns></returns>
		string StringifyNumber(double number);
	}
}