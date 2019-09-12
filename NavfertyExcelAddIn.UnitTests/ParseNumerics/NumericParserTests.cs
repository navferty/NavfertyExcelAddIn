using System;
using System.Reflection;
using System.Linq.Expressions;
using Moq;
using NavfertyExcelAddIn.ParseNumerics;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Microsoft.Office.Interop.Excel;
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

            var ws = new Mock<Worksheet>(MockBehavior.Strict);
            ws.Setup(x => x.Name).Returns("WsName");
            selection.Setup(x => x.Worksheet).Returns(ws.Object);

            var areas = new Mock<Areas>(MockBehavior.Strict);
            areas.Setup(x => x.GetEnumerator()).Returns(new[] { selection.Object }.GetEnumerator());
            selection.Setup(x => x.Areas).Returns(areas.Object);

            NumericParser.Parse(selection.Object);
        }

        private bool VerifyParsed(object[,] parsedValues)
        {
            var expected = new object[,] { { 1, 1m, "abc" }, { 123.123m, 321.321m, null } };
            CollectionAssert.AreEqual(expected, parsedValues);
            return true;
        }
    }
}
