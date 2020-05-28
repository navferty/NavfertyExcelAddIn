namespace NavfertyExcelAddIn.Transliterate
{
	/// <summary>
	/// Replace all cyrillic letters, that have matching analogues in latin.
	/// Not matched letters will not be replaced. <br/> <br/>
	/// Examples: <br/>
	/// "абвгдек" -> "aбbгдek" (a,b,e,k are latin, others - cyrillic) <br/>
	/// "АБВГДЕК" -> "AБBГДEK" (A,B,E,K are latin, others - cyrillic) <br/>
	/// </summary>
	public interface ICyrillicLettersReplacer
	{
		string ReplaceCyrillicCharsWithLatin(string input);
	}
}