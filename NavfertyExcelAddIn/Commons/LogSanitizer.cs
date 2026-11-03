using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace NavfertyExcelAddIn.Commons
{
	/// <summary>
	/// Masks user data (cell values, workbook/worksheet names, file paths) before it is written to logs,
	/// so log files can be safely attached to public bug reports.
	/// </summary>
	internal static class LogSanitizer
	{
		// Random per process, so short values can't be recovered from the hash by a dictionary attack.
		private static readonly byte[] Salt = GenerateSalt();

		/// <summary>Cell value or any other user-provided data.</summary>
		public static string Value(object? value)
			=> value is null ? "null" : Mask(Convert.ToString(value, CultureInfo.InvariantCulture));

		/// <summary>Workbook, worksheet, table or column name.</summary>
		public static string Name(string? name) => Mask(name);

		/// <summary>File path; only the extension is kept.</summary>
		public static string Path(string? path)
		{
			if (path is null)
				return "null";

			string extension;
			try
			{
				extension = System.IO.Path.GetExtension(path);
			}
			catch (ArgumentException)
			{
				extension = string.Empty;
			}

			return $"<path ext={extension}, id={Hash(path)}>";
		}

		private static string Mask(string? text)
		{
			if (text is null)
				return "null";

			return text.Length == 0 ? "<empty>" : $"<len={text.Length}, id={Hash(text)}>";
		}

		private static string Hash(string text)
		{
			using (var sha = SHA256.Create())
			{
				var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(text + Convert.ToBase64String(Salt)));
				return BitConverter.ToString(bytes, 0, 4).Replace("-", string.Empty);
			}
		}

		private static byte[] GenerateSalt()
		{
			var salt = new byte[16];
			using (var rng = RandomNumberGenerator.Create())
			{
				rng.GetBytes(salt);
			}
			return salt;
		}
	}
}
