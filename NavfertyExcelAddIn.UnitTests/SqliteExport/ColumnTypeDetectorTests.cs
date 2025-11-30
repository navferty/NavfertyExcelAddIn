using NavfertyExcelAddIn.SqliteExport;

namespace NavfertyExcelAddIn.UnitTests.SqliteExport;

public class ColumnTypeDetectorTests
{
	[Test]
	public async Task DetectColumnType_AllIntegers_ReturnsInteger()
	{
		var values = new object[4, 2];
		values[1, 1] = 1;
		values[2, 1] = 2;
		values[3, 1] = 3;

		var result = ColumnTypeDetector.DetectColumnType(values, 1, 1, 3);

		await Assert.That(result).IsEqualTo(ColumnTypeDetector.SqliteColumnType.INTEGER);
	}

	[Test]
	public async Task DetectColumnType_AllFloats_ReturnsReal()
	{
		var values = new object[4, 2];
		values[1, 1] = 1.5;
		values[2, 1] = 2.7;
		values[3, 1] = 3.14;

		var result = ColumnTypeDetector.DetectColumnType(values, 1, 1, 3);

		await Assert.That(result).IsEqualTo(ColumnTypeDetector.SqliteColumnType.REAL);
	}

	[Test]
	public async Task DetectColumnType_AllText_ReturnsText()
	{
		var values = new object[4, 2];
		values[1, 1] = "apple";
		values[2, 1] = "banana";
		values[3, 1] = "cherry";

		var result = ColumnTypeDetector.DetectColumnType(values, 1, 1, 3);

		await Assert.That(result).IsEqualTo(ColumnTypeDetector.SqliteColumnType.TEXT);
	}

	[Test]
	public async Task DetectColumnType_AllDates_ReturnsNumeric()
	{
		var values = new object[4, 2];
		values[1, 1] = new DateTime(2023, 1, 1);
		values[2, 1] = new DateTime(2023, 2, 1);
		values[3, 1] = new DateTime(2023, 3, 1);

		var result = ColumnTypeDetector.DetectColumnType(values, 1, 1, 3);

		await Assert.That(result).IsEqualTo(ColumnTypeDetector.SqliteColumnType.NUMERIC);
	}

	[Test]
	public async Task DetectColumnType_AllBooleans_ReturnsInteger()
	{
		var values = new object[4, 2];
		values[1, 1] = true;
		values[2, 1] = false;
		values[3, 1] = true;

		var result = ColumnTypeDetector.DetectColumnType(values, 1, 1, 3);

		await Assert.That(result).IsEqualTo(ColumnTypeDetector.SqliteColumnType.INTEGER);
	}

	[Test]
	public async Task DetectColumnType_AllBlobs_ReturnsBlob()
	{
		var values = new object[4, 2];
		values[1, 1] = new byte[] { 1, 2, 3 };
		values[2, 1] = new byte[] { 4, 5, 6 };
		values[3, 1] = new byte[] { 7, 8, 9 };

		var result = ColumnTypeDetector.DetectColumnType(values, 1, 1, 3);

		await Assert.That(result).IsEqualTo(ColumnTypeDetector.SqliteColumnType.BLOB);
	}

	[Test]
	public async Task DetectColumnType_MixedTypes_ReturnsText()
	{
		var values = new object[4, 2];
		values[1, 1] = 1;
		values[2, 1] = "text";
		values[3, 1] = 3.14;

		var result = ColumnTypeDetector.DetectColumnType(values, 1, 1, 3);

		await Assert.That(result).IsEqualTo(ColumnTypeDetector.SqliteColumnType.TEXT);
	}

	[Test]
	public async Task DetectColumnType_IntegersWithNulls_ReturnsInteger()
	{
		var values = new object?[5, 2];
		values[1, 1] = 1;
		values[2, 1] = null;
		values[3, 1] = 3;
		values[4, 1] = "";

		var result = ColumnTypeDetector.DetectColumnType(values, 1, 1, 4);

		await Assert.That(result).IsEqualTo(ColumnTypeDetector.SqliteColumnType.INTEGER);
	}

	[Test]
	public async Task DetectColumnType_EmptyColumn_ReturnsText()
	{
		var values = new object?[4, 2];
		values[1, 1] = null;
		values[2, 1] = "";
		values[3, 1] = null;

		var result = ColumnTypeDetector.DetectColumnType(values, 1, 1, 3);

		await Assert.That(result).IsEqualTo(ColumnTypeDetector.SqliteColumnType.TEXT);
	}

	[Test]
	public async Task DetectColumnType_DoubleAsInteger_ReturnsInteger()
	{
		var values = new object[4, 2];
		values[1, 1] = 1.0;
		values[2, 1] = 2.0;
		values[3, 1] = 3.0;

		var result = ColumnTypeDetector.DetectColumnType(values, 1, 1, 3);

		await Assert.That(result).IsEqualTo(ColumnTypeDetector.SqliteColumnType.INTEGER);
	}

	[Test]
	public async Task DetectColumnType_StringNumbers_ReturnsInteger()
	{
		var values = new object[4, 2];
		values[1, 1] = "1";
		values[2, 1] = "2";
		values[3, 1] = "3";

		var result = ColumnTypeDetector.DetectColumnType(values, 1, 1, 3);

		await Assert.That(result).IsEqualTo(ColumnTypeDetector.SqliteColumnType.INTEGER);
	}

	[Test]
	public async Task DetectColumnType_StringFloats_ReturnsReal()
	{
		var values = new object[4, 2];
		values[1, 1] = "1.5";
		values[2, 1] = "2.7";
		values[3, 1] = "3.14";

		var result = ColumnTypeDetector.DetectColumnType(values, 1, 1, 3);

		await Assert.That(result).IsEqualTo(ColumnTypeDetector.SqliteColumnType.REAL);
	}

	[Test]
	public async Task DetectColumnType_NullValues_ReturnsText()
	{
		var result = ColumnTypeDetector.DetectColumnType(null, 1, 1, 3);

		await Assert.That(result).IsEqualTo(ColumnTypeDetector.SqliteColumnType.TEXT);
	}
}
