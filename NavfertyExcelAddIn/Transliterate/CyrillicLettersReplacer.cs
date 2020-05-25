using System.Collections.Generic;
using System.Linq;

namespace NavfertyExcelAddIn.Transliterate
{
	public class CyrillicLettersReplacer : ICyrillicLettersReplacer
	{
		private readonly Dictionary<char, char> cyrillicToLatinChars = new Dictionary<char, char>
		{
			// Uppercase modern Cyrillic characters
			{ '\u0410', 'A' },
			{ '\u0412', 'B' },
			{ '\u0415', 'E' },
			{ '\u041A', 'K' },
			{ '\u041C', 'M' },
			{ '\u041D', 'H' },
			{ '\u041E', 'O' },
			{ '\u0420', 'P' },
			{ '\u0421', 'C' },
			{ '\u0422', 'T' },
			{ '\u0423', 'Y' },
			{ '\u0425', 'X' },

			// Lowercase modern Cyrillic characters
			{ '\u0430', 'a' },
			{ '\u0432', 'b' },
			{ '\u0435', 'e' },
			{ '\u043A', 'k' },
			{ '\u043C', 'm' },
			{ '\u043D', 'h' },
			{ '\u043E', 'o' },
			{ '\u0440', 'p' },
			{ '\u0441', 'c' },
			{ '\u0442', 't' },
			{ '\u0443', 'y' },
			{ '\u0445', 'x' }
		};

		public string ReplaceCyrillicCharsWithLatin(string input)
		{
			var replaced = input.ToCharArray()
				.Select(x =>
				{
					if (cyrillicToLatinChars.TryGetValue(x, out var found))
						return found;
					return x;
				}).ToArray();
			return new string(replaced);
		}
	}
}
