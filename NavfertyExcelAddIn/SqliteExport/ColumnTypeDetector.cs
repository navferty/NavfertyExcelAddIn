using System;

namespace NavfertyExcelAddIn.SqliteExport
{
	internal class ColumnTypeDetector
	{
		public enum SqliteColumnType
		{
			INTEGER,
			REAL,
			TEXT,
			NUMERIC,  // For dates and times
			BLOB
		}

		private const double FloatingPointTolerance = 1e-10;

		public static SqliteColumnType DetectColumnType(object[,] values, int colIndex, int startRow, int endRow)
		{
			if (values == null)
			{
				return SqliteColumnType.TEXT;
			}

			bool hasValues = false;
			bool allIntegers = true;
			bool allNumeric = true;
			bool allDates = true;
			bool allBooleans = true;

			for (int row = startRow; row <= endRow; row++)
			{
				var value = values[row, colIndex];
				if (value == null || (value is string str && string.IsNullOrWhiteSpace(str)))
				{
					continue;
				}

				hasValues = true;

				// Check for boolean values
				if (!(value is bool))
				{
					allBooleans = false;
				}

				// Check for DateTime values
				if (!(value is DateTime))
				{
					allDates = false;
				}

				// Check if the value is a number
				if (value is double || value is float || value is decimal || 
				    value is int || value is long || value is short || value is byte)
				{
					// Check if it's an integer value
					if (value is double dbl)
					{
						if (Math.Abs(dbl - Math.Round(dbl)) > FloatingPointTolerance)
						{
							allIntegers = false;
						}
					}
					else if (value is float flt)
					{
						if (Math.Abs(flt - Math.Round(flt)) > FloatingPointTolerance)
						{
							allIntegers = false;
						}
					}
					else if (value is decimal dec)
					{
						if (dec != Math.Round(dec))
						{
							allIntegers = false;
						}
					}
					// else: int, long, short, byte are already integers
				}
				else if (value is DateTime || value is bool)
				{
					// DateTime and bool are not considered numeric for REAL/INTEGER type
					allNumeric = false;
					allIntegers = false;
				}
				else if (value is byte[])
				{
					// Binary data
					allNumeric = false;
					allIntegers = false;
					allDates = false;
					allBooleans = false;
				}
				else
				{
					// Try to parse as a number
					if (double.TryParse(value.ToString(), out double parsed))
					{
						if (Math.Abs(parsed - Math.Round(parsed)) > FloatingPointTolerance)
						{
							allIntegers = false;
						}
					}
					else
					{
						// Not a number at all
						allNumeric = false;
						allIntegers = false;
					}
				}
			}

			if (!hasValues)
			{
				return SqliteColumnType.TEXT;
			}

			// Prioritize specific types
			if (allBooleans)
			{
				return SqliteColumnType.INTEGER; // SQLite stores booleans as INTEGER (0 or 1)
			}

			if (allDates)
			{
				return SqliteColumnType.NUMERIC; // SQLite NUMERIC type for dates
			}

			if (allNumeric)
			{
				return allIntegers ? SqliteColumnType.INTEGER : SqliteColumnType.REAL;
			}

			// Check if all values are byte arrays (binary data)
			bool allBlobs = true;
			for (int row = startRow; row <= endRow; row++)
			{
				var value = values[row, colIndex];
				if (value != null && !(value is byte[]) && !(value is string str && string.IsNullOrWhiteSpace(str)))
				{
					allBlobs = false;
					break;
				}
			}

			if (allBlobs)
			{
				return SqliteColumnType.BLOB;
			}

			return SqliteColumnType.TEXT;
		}
	}
}
