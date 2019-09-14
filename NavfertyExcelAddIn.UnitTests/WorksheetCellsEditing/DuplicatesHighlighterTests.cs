using System.Collections;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using Moq;
using NavfertyExcelAddIn.WorksheetCellsEditing;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace NavfertyExcelAddIn.UnitTests.WorksheetCellsEditing
{
    [TestClass]
    public class DuplicatesHighlighterTests : TestsBase
    {
        private Mock<Range> selection;

        private DuplicatesHighlighter duplicatesHighlighter;

        [TestInitialize]
        public void BeforeEachTest()
        {
            selection = GetRangeStub();

            SetRangeExtentionsStub();

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

            Assert.AreEqual(3, RangeExtensionsStub.SetColorInvocations.Count());
        }

        [TestMethod]
        public void HighlightDuplicates_ManyItemsWithSameColors_Success()
        {
            var cells = Enumerable.Range(1, 1000).Select(x => GetCellStub((x % 100).ToString()));
            selection.As<IEnumerable>().Setup(x => x.GetEnumerator()).Returns(cells.Select(c => c.Object).GetEnumerator());

            duplicatesHighlighter.HighlightDuplicates(selection.Object);

            // all 57 colors
            Assert.AreEqual(57, RangeExtensionsStub.SetColorInvocations.Count());
        }
    }
}
