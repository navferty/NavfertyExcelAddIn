using System.Linq.Expressions;
using System.Reflection;

using Microsoft.Office.Interop.Excel;

using Moq;

using NavfertyExcelAddIn.ParseNumerics;

using Range = Microsoft.Office.Interop.Excel.Range;

namespace NavfertyExcelAddIn.UnitTests.ParseNumerics;

public class NumericParserTests
{
	[Test]
	public void VariousValuesInSelection_ParseNumerics_ValuesParsedSuccessfully()
	{
		var selection = new Mock<Range>(MockBehavior.Strict);
		var values = new object?[,] { { 1, "1", "abc" }, { "123.123", "321 , 321", null } };

		selection.SetupGet(x => x.get_Value(Missing.Value)).Returns(values);
		selection.SetupSet(x => x.Value = It.Is<object[,]>(z => VerifyParsed(z).GetAwaiter().GetResult()));
		Expression<Func<Range, string>> getAddress = x => x.get_Address(It.IsAny<object>(), It.IsAny<object>(),
							It.IsAny<XlReferenceStyle>(), It.IsAny<object>(), It.IsAny<object>());
		selection.SetupGet(getAddress).Returns("Address");
		selection.SetupGet(x => x.Row).Returns(1);
		selection.SetupGet(x => x.Column).Returns(1);

		var ws = new Mock<Worksheet>(MockBehavior.Strict);
		var wb = new Mock<Workbook>(MockBehavior.Strict);
		ws.Setup(x => x.Name).Returns("WsName");
		wb.Setup(x => x.Name).Returns("WbName");
		selection.Setup(x => x.Worksheet).Returns(ws.Object);
		ws.Setup(x => x.Parent).Returns(wb.Object);

		var areas = new Mock<Areas>(MockBehavior.Strict);
		areas.Setup(x => x.GetEnumerator()).Returns(new[] { selection.Object }.GetEnumerator());
		selection.Setup(x => x.Areas).Returns(areas.Object);
		var numericParser = new NumericParserService();

		numericParser.Parse(selection.Object);
	}

	private async Task<bool> VerifyParsed(object?[,] parsedValues)
	{
		var expected = new object?[,] { { 1, 1d, "abc" }, { 123.123d, 321.321d, null } };
		await Assert.That(parsedValues).IsEquivalentTo(expected);
		return true;
	}
}
