using System;
using System.Linq.Expressions;
using System.Reflection;

using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using Moq;

using NavfertyExcelAddIn.ParseNumerics;

using Range = Microsoft.Office.Interop.Excel.Range;

namespace NavfertyExcelAddIn.UnitTests.ParseNumerics
{
	[TestClass]
	public class NumericParserTests
	{
		private NumericParser NumericParser;

		[TestInitialize]
		public void BeforeEachTest()
		{
			NumericParser = new NumericParser();
		}

		[TestMethod]
		public void ParsedSuccessfully() // TODO naming
		{
			var selection = new Mock<Range>(MockBehavior.Strict);
			var values = new object[,] { { 1, "1", "abc" }, { "123.123", "321 , 321", null } };

			selection.SetupGet(x => x.get_Value(Missing.Value)).Returns(values);
			selection.SetupSet(x => x.Value = It.Is<object[,]>(z => VerifyParsed(z)));
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


			var cells = new Mock<Range>(MockBehavior.Default);
			cells.Setup(x => x.GetEnumerator()).Returns(new[] { selection.Object }.GetEnumerator());
			//selection.SetupSet(x => x.get_Cells = It.Is<object[,]>(z => VerifyParsed(z)));
			//selection.Setup(x => x.Cells).Returns(cells.Object.GetEnumerator);

			//areas.Setup(x => x.ce .Cells).Returns(cells.Object);

			NumericParser.Parse(selection.Object);
		}

		private bool VerifyParsed(object[,] parsedValues)
		{
			var expected = new object[,] { { 1, 1d, "abc" }, { 123.123d, 321.321d, null } };
			CollectionAssert.AreEqual(expected, parsedValues);
			return true;
		}
	}
}
