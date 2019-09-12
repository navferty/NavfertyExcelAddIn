using System.Collections;
using System.Reflection;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using NavfertyExcelAddIn.WorksheetCellsEditing;

using Range = Microsoft.Office.Interop.Excel.Range;

namespace NavfertyExcelAddIn.UnitTests.WorksheetCellsEditing
{
    [TestClass]
    public class CellsToMarkdownReaderTests : ExcelTests
    {
        private Mock<Range> selection;
        private Mock<Range> cell;

        private CellsToMarkdownReader cellsToMarkdownReader;

        [TestInitialize]
        public override void BeforeEachTest()
        {
            base.BeforeEachTest();

            selection = GetRangeStub();
            cell = new Mock<Range>();

            var rows = new Mock<Range>();
            rows.As<IEnumerable>()
                .Setup(x => x.GetEnumerator())
                .Returns(new [] { rows.Object, rows.Object, rows.Object }.GetEnumerator());

            var cells = new Mock<Range>();
            cells.As<IEnumerable>()
                .SetupSequence(x => x.GetEnumerator())
                .Returns(new [] { cell.Object, cell.Object, cell.Object }.GetEnumerator())
                .Returns(new [] { cell.Object, cell.Object, cell.Object }.GetEnumerator())
                .Returns(new [] { cell.Object, cell.Object, cell.Object }.GetEnumerator());

            selection.Setup(x => x.Rows).Returns(rows.Object);
            rows.Setup(x => x.Cells).Returns(cells.Object);

            var columns = new Mock<Range>();
            columns.Setup(x => x.Count).Returns(3);
            selection.Setup(x => x.Columns).Returns(columns.Object);

            cellsToMarkdownReader = new CellsToMarkdownReader();
        }

        [TestMethod]
        public void ReadTableAsMarkdown()
        {
            var setup = cell.SetupSequence(x => x.get_Value(Missing.Value));
            for (var i = 0; i < 9; i++)
            {
                setup = setup.Returns(new string((char)('a' + i), 3));
            }

            var result = cellsToMarkdownReader.ReadToMarkdown(selection.Object);

            Assert.AreEqual(
@"|aaa|bbb|ccc|
|---|---|---|
|ddd|eee|fff|
|ggg|hhh|iii|
", result);
        }
    }
}
