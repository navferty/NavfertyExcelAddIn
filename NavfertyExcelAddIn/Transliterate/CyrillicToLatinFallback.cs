using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NavfertyExcelAddIn.Transliterate
{
	// based on https://github.com/dotnet/samples/blob/master/core/encoding/cyrillic-to-latin

	public class CyrillicToLatinFallback : EncoderFallback
	{
		// transliteration standard: ICAO Doc 9303 (and Russian passport standard, Приказ МИД №4271)
		// https://www.icao.int/publications/Documents/9303_p3_cons_en.pdf
		private readonly Dictionary<char, string> table = new Dictionary<char, string>
			{
				// Uppercase modern Cyrillic characters
				{ '\u0410', "A" },
				{ '\u0411', "B" },
				{ '\u0412', "V" },
				{ '\u0413', "G" },
				{ '\u0414', "D" },
				{ '\u0415', "E" },
				{ '\u0401', "E" }, // Ё
				{ '\u0416', "Zh" },
				{ '\u0417', "Z" },
				{ '\u0418', "I" },
				{ '\u0419', "I" },
				{ '\u041A', "K" },
				{ '\u041B', "L" },
				{ '\u041C', "M" },
				{ '\u041D', "N" },
				{ '\u041E', "O" },
				{ '\u041F', "P" },
				{ '\u0420', "R" },
				{ '\u0421', "S" },
				{ '\u0422', "T" },
				{ '\u0423', "U" },
				{ '\u0424', "F" },
				{ '\u0425', "Kh" },
				{ '\u0426', "Ts" },
				{ '\u0427', "Ch" },
				{ '\u0428', "Sh" },
				{ '\u0429', "Shch" },
				{ '\u042A', "Ie" },    // Hard sign
				{ '\u042B', "Y" },
				{ '\u042C', "" },    // Soft sign not transliterated
				{ '\u042D', "E" },
				{ '\u042E', "Iu" },
				{ '\u042F', "Ia" },

				// Lowercase modern Cyrillic characters
				{ '\u0430', "a" },
				{ '\u0431', "b" },
				{ '\u0432', "v" },
				{ '\u0433', "g" },
				{ '\u0434', "d" },
				{ '\u0435', "e" },
				{ '\u0451', "e" }, // ё
				{ '\u0436', "zh" },
				{ '\u0437', "z" },
				{ '\u0438', "i" },
				{ '\u0439', "i" },
				{ '\u043A', "k" },
				{ '\u043B', "l" },
				{ '\u043C', "m" },
				{ '\u043D', "n" },
				{ '\u043E', "o" },
				{ '\u043F', "p" },
				{ '\u0440', "r" },
				{ '\u0441', "s" },
				{ '\u0442', "t" },
				{ '\u0443', "u" },
				{ '\u0444', "f" },
				{ '\u0445', "kh" },
				{ '\u0446', "ts" },
				{ '\u0447', "ch" },
				{ '\u0448', "sh" },
				{ '\u0449', "shch" },
				{ '\u044A', "ie" },   // Hard sign
				{ '\u044B', "y" },
				{ '\u044C', "" },   // Soft sign not transliterated
				{ '\u044D', "e" },
				{ '\u044E', "iu" },
				{ '\u044F', "ia" }
			};

		public override EncoderFallbackBuffer CreateFallbackBuffer()
		{
			return new CyrillicToLatinFallbackBuffer(table);
		}

		public override int MaxCharCount
		{
			// Maximum is "Shch" and "shch"
			get { return table.Max(x => x.Value.Length); }
		}
	}

	public class CyrillicToLatinFallbackBuffer : EncoderFallbackBuffer
	{
		private readonly Dictionary<char, string> table;

		private int bufferIndex = -1;
		private int leftToReturn = -1;
		private string buffer = string.Empty;

		internal CyrillicToLatinFallbackBuffer(Dictionary<char, string> table)
		{
			this.table = table;
		}

		public override bool Fallback(char charUnknownHigh, char charUnknownLow, int index)
		{
			// There's no need to handle surrogates.
			return false;
		}

		public override bool Fallback(char charUnknown, int index)
		{
			if (!table.ContainsKey(charUnknown))
				return false;

			buffer = table[charUnknown];
			leftToReturn = buffer.Length - 1;
			bufferIndex = -1;
			return true;
		}

		public override char GetNextChar()
		{
			if (leftToReturn < 0)
				return '\u0000';

			leftToReturn--;
			bufferIndex++;
			return buffer[bufferIndex];
		}

		public override bool MovePrevious()
		{
			if (bufferIndex > 0)
			{
				bufferIndex--;
				leftToReturn++;
				return true;
			}
			return false;
		}

		public override int Remaining => leftToReturn;
	}
}
