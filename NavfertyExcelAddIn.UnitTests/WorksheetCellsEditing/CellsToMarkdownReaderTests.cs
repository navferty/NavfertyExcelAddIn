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

        private CellsToMarkdownReader cellsToMarkdownReader;

        [TestInitialize]
        public override void BeforeEachTest()
        {
            base.BeforeEachTest();

            selection = new Mock<Range>();

            selection.Setup(x => x.Columns).Returns(selection.Object);
            selection.Setup(x => x.Count).Returns(3);

            selection.Setup(x => x.Rows).Returns(selection.Object);
            selection.Setup(x => x.Cells).Returns(selection.Object);
            selection.Setup(x => x.GetEnumerator()).Returns(GetRangeEnumerator);

            cellsToMarkdownReader = new CellsToMarkdownReader();
        }

        [TestMethod]
        public void ReadTableAsMarkdown()
        {
            var setup = selection.SetupSequence(x => x.get_Value(Missing.Value));
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

        private IEnumerator GetRangeEnumerator()
        {
            yield return selection.Object;
            yield return selection.Object;
            yield return selection.Object;
        }
    }
}
