using System.Reflection;

using NavfertyExcelAddIn.SqliteImport;

namespace NavfertyExcelAddIn.UnitTests.SqliteImport;

public class SqliteSheetNameTests
{
	[Test]
	public async Task HandlesReservedNamesTruncationAndCollisions()
	{
		var method = typeof(SqliteImporter).GetMethod("GetSheetName", BindingFlags.NonPublic | BindingFlags.Static)!;
		var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		string GetName(string value) => (string)method.Invoke(null, new object[] { value, names })!;
		await Assert.That(GetName("History")).IsEqualTo("History_");
		await Assert.That(GetName("history")).IsEqualTo("history_ (2)");
		var longName = new string('a', 30) + "'long";
		await Assert.That(GetName(longName)).IsEqualTo(new string('a', 30));
		var duplicate = GetName(longName);
		await Assert.That(duplicate).IsEqualTo(new string('a', 27) + " (2)");
		await Assert.That(GetName("'''")).IsEqualTo("Sheet");
	}
}
