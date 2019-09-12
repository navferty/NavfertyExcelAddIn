using System.Collections;
using System.Reflection;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

using NavfertyExcelAddIn.WorksheetCellsEditing;
using NavfertyExcelAddIn.FindFormulaErrors;

using Range = Microsoft.Office.Interop.Excel.Range;

namespace NavfertyExcelAddIn.UnitTests.WorksheetCellsEditing
{
    [TestClass]
    public class CaseTogglerTests : ExcelTests
    {
        private Mock<Range> selection;

        private CaseToggler caseToggler;

        [TestInitialize]
        public override void BeforeEachTest()
        {
            base.BeforeEachTest();

            selection = GetRangeStub();
            selection.Setup(x => x.set_Value(It.IsAny<object>(), It.IsAny<object>()));

            caseToggler = new CaseToggler();
        }

        [TestMethod]
        public void ToggleCase_SingleCell_Converted()
        {
            var value = "abc";
            selection.As<IEnumerable>().Setup(x => x.GetEnumerator()).Returns(new[] { GetCellStub("abc").Object }.GetEnumerator());
            selection.Setup(x => x.get_Value(Missing.Value)).Returns(value);

            caseToggler.ToggleCase(selection.Object);


            var expected = "ABC";
            selection.Verify(x => x.set_Value(It.IsAny<object>(), It.IsAny<string>()), Times.Once);
            selection.Verify(x => x.set_Value(It.IsAny<object>(), It.Is<string>(v => expected == v)));
        }

        [TestMethod]
        public void ToggleCase_AllItemsConverted()
        {
            var values = new object[,]
            {
                { "abc", "def", "ghi" },
                { "jkl", "123", 123d },
                { "", null, (int)CVErrEnum.ErrNA }
            };
            selection.As<IEnumerable>().Setup(x => x.GetEnumerator()).Returns(new[] { GetCellStub("abc").Object }.GetEnumerator());
            selection.Setup(x => x.get_Value(Missing.Value)).Returns(values);

            caseToggler.ToggleCase(selection.Object);


            var expected = new object[,]
            {
                { "ABC", "DEF", "GHI" },
                { "JKL", "123", 123d },
                { "", null, (int)CVErrEnum.ErrNA }
            };
            selection.Verify(x => x.set_Value(It.IsAny<object>(), It.IsAny<object[,]>()), Times.Once);
            selection.Verify(x => x.set_Value(It.IsAny<object>(), It.Is<object[,]>(v => AssertAssignedValue(expected, v))));
        }
        private bool AssertAssignedValue(object[,] expected, object[,] value)
        {
            CollectionAssert.AreEquivalent(expected, value);
            return true;
        }

        [TestMethod]
        [DataRow("abc", "ABC")]
        [DataRow("ABC", "Abc")]
        [DataRow("Abc", "abc")]

        [DataRow("A", "a")]
        [DataRow("a", "A")]
        public void ToggleCase_ConvertedRightOrder(string initial, string result)
        {
            selection.As<IEnumerable>().Setup(x => x.GetEnumerator()).Returns(new[] { GetCellStub(initial).Object }.GetEnumerator());
            selection.Setup(x => x.get_Value(Missing.Value)).Returns(initial);

            caseToggler.ToggleCase(selection.Object);


            selection.Verify(x => x.set_Value(It.IsAny<object>(), It.IsAny<string>()), Times.Once);
            selection.Verify(x => x.set_Value(It.IsAny<object>(), It.Is<string>(v => result == v)));
        }

        [TestMethod]
        public void ToggleCase_NoCellsWithLetters_NoChanges()
        {
            var values = new object[,]
            {
                { "", "123", 123d },
                { "", null, (int)CVErrEnum.ErrNA }
            };
            selection.As<IEnumerable>().Setup(x => x.GetEnumerator()).Returns(
                new[]
                {
                    GetCellStub("").Object, GetCellStub("123").Object, GetCellStub(123d).Object,
                    GetCellStub("").Object, GetCellStub(null).Object, GetCellStub((int)CVErrEnum.ErrNA).Object
                }.GetEnumerator());

            selection.Setup(x => x.get_Value(Missing.Value)).Returns(values);


            caseToggler.ToggleCase(selection.Object);

            selection.Verify(x => x.set_Value(It.IsAny<object>(), It.IsAny<object[,]>()), Times.Never);
        }
    }
}
