using System.Text;

namespace NavfertyExcelAddIn.Transliterate
{
	public class Transliterator : ITransliterator
	{
		public string Transliterate(string input)
		{
			if (string.IsNullOrWhiteSpace(input))
				return input;

			// Instantiate the encoder
			var encoding = Encoding.GetEncoding(
				name: "us-ascii",
				encoderFallback: new CyrillicToLatinFallback(),
				decoderFallback: new DecoderExceptionFallback());

			var encoder = encoding.GetEncoder();
			var decoder = encoding.GetDecoder();

			var inputChars = input.ToCharArray();

			// Encode characters
			var byteCount = encoder.GetByteCount(
				chars: inputChars,
				index: 0,
				count: inputChars.Length,
				flush: false);

			var bytes = new byte[byteCount];
			var bytesWritten = encoder.GetBytes(
				chars: inputChars,
				charIndex: 0,
				charCount: inputChars.Length,
				bytes: bytes,
				byteIndex: 0,
				flush: false);

			// Decode characters back to Unicode
			char[] charsToWrite = new char[decoder.GetCharCount(bytes, 0, byteCount)];
			decoder.GetChars(
				bytes: bytes,
				byteIndex: 0,
				byteCount: bytesWritten,
				chars: charsToWrite,
				charIndex: 0);

			return new string(charsToWrite);
		}
	}
}
