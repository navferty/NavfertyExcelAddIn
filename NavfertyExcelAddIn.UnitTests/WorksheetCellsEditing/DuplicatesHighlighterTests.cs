using System.Collections;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using Moq;
using NavfertyExcelAddIn.WorksheetCellsEditing;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace NavfertyExcelAddIn.UnitTests.WorksheetCellsEditing
{
    [TestClass]
    public class DuplicatesHighlighterTests : ExcelTests
    {
        private Mock<Range> selection;
        private Mock<Range> union;

        private DuplicatesHighlighter duplicatesHighlighter;

        [TestInitialize]
        public override void BeforeEachTest()
        {
            base.BeforeEachTest();

            selection = GetRangeStub();
            union = GetRangeStub();

            rangeExtensions.Setup(x => x.Union(It.IsAny<Range>(), It.IsAny<Range>())).Returns(union.Object);

            duplicatesHighlighter = new DuplicatesHighlighter();
        }

        [TestMethod]
        public void HighlightDuplicates_ThreeItems_Success()
        {
            var cells = new[]
            {
                GetCellStub("a"),
                GetCellStub("b"),
                GetCellStub("c"),
                GetCellStub("a"),
                GetCellStub("b"),
                GetCellStub("qwe"),
                GetCellStub(2d),
                GetCellStub(3d),
                GetCellStub(2d)
            };
            selection.As<IEnumerable>().Setup(x => x.GetEnumerator()).Returns(cells.Select(c => c.Object).GetEnumerator());

            duplicatesHighlighter.HighlightDuplicates(selection.Object);

            rangeExtensions.Verify(x => x.SetColor(It.IsAny<Range>(), It.IsAny<int>()), Times.Exactly(3));
        }

        [TestMethod]
        public void HighlightDuplicates_ManyItemsWithSameColors_Success()
        {
            var cells = Enumerable.Range(1, 1000).Select(x => GetCellStub((x % 100).ToString()));
            selection.As<IEnumerable>().Setup(x => x.GetEnumerator()).Returns(cells.Select(c => c.Object).GetEnumerator());

            duplicatesHighlighter.HighlightDuplicates(selection.Object);

            rangeExtensions.Verify(x => x.SetColor(It.IsAny<Range>(), It.IsAny<int>()), Times.Exactly(57)); // all 57 colors
        }
    }
}
