using System.Reflection;
using Moq;
using NavfertyExcelAddIn.ParseNumerics;

using Microsoft.VisualStudio.TestTools.UnitTesting;

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
