using System.Collections;
using System.Reflection;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using NavfertyExcelAddIn.FindFormulaErrors;
using NavfertyExcelAddIn.WorksheetCellsEditing;

using Range = Microsoft.Office.Interop.Excel.Range;

namespace NavfertyExcelAddIn.UnitTests.WorksheetCellsEditing
{
    [TestClass]
    public class EmptySpaceTrimmerTests : TestsBase
    {
        private Mock<Range> selection;

        private EmptySpaceTrimmer emptySpaceTrimmer;

        [TestInitialize]
        public void BeforeEachTest()
        {
            selection = GetRangeStub();
            selection.Setup(x => x.set_Value(It.IsAny<object>(), It.IsAny<object>()));

            emptySpaceTrimmer = new EmptySpaceTrimmer();
        }

        [TestMethod]
        public void TrimSpaces_SingleCell_Converted()
        {
            var value = "a   \r\nb     c\t";
            selection.As<IEnumerable>().Setup(x => x.GetEnumerator()).Returns(new[] { GetCellStub("abc").Object }.GetEnumerator());
            selection.Setup(x => x.get_Value(Missing.Value)).Returns(value);

            emptySpaceTrimmer.TrimSpaces(selection.Object);


            var expected = "a b c";
            selection.Verify(x => x.set_Value(It.IsAny<object>(), It.IsAny<string>()), Times.Once);
            selection.Verify(x => x.set_Value(It.IsAny<object>(), It.Is<string>(v => expected == v)));
        }

        [TestMethod]
        public void TrimSpaces_AllItemsTrimmed()
        {
            var values = new object[,]
            {
                { "abc ", " def", "  ghi\t" },
                { "jk\r\nl", "   \r\n   \t", "   " },
                { "", null, (int)CVErrEnum.ErrNA }
            };
            selection.Setup(x => x.get_Value(Missing.Value)).Returns(values);


            emptySpaceTrimmer.TrimSpaces(selection.Object);

            var expected = new object[,]
            {
                { "abc", "def", "ghi" },
                { "jk l", null, null },
                { null, null, (int)CVErrEnum.ErrNA }
            };
            selection.Verify(x => x.set_Value(It.IsAny<object>(), It.IsAny<object[,]>()), Times.Once);
            selection.Verify(x => x.set_Value(It.IsAny<object>(), It.Is<object[,]>(v => AssertAssignedValue(expected, v))));
        }
        private bool AssertAssignedValue(object[,] expected, object[,] value)
        {
            CollectionAssert.AreEquivalent(expected, value);
            return true;
        }
    }
}
