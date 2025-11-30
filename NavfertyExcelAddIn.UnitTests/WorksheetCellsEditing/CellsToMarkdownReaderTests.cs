using System.Collections;
using System.Reflection;

using Moq;

using NavfertyExcelAddIn.WorksheetCellsEditing;

using Range = Microsoft.Office.Interop.Excel.Range;

namespace NavfertyExcelAddIn.UnitTests.WorksheetCellsEditing
{
	public class CellsToMarkdownReaderTests : TestsBase
	{
		private Mock<Range> selection = null!;

		[Before(HookType.Test)]
		public void BeforeEachTest()
		{
			selection = new Mock<Range>();

			selection.Setup(x => x.Columns).Returns(selection.Object);
			selection.Setup(x => x.Count).Returns(3);

			selection.Setup(x => x.Rows).Returns(selection.Object);
			selection.Setup(x => x.Cells).Returns(selection.Object);
			selection.Setup(x => x.GetEnumerator()).Returns(GetRangeEnumerator);
		}

		[Test]
		public async Task ReadTableAsMarkdown()
		{
			var setup = selection.SetupSequence(x => x.get_Value(Missing.Value));
			for (var i = 0; i < 9; i++)
			{
				setup = setup.Returns(new string((char)('a' + i), 3));
			}
			var cellsToMarkdownReader = new CellsToMarkdownReader();

			var result = cellsToMarkdownReader.ReadToMarkdown(selection.Object);

			var expected = """
				|aaa|bbb|ccc|
				|---|---|---|
				|ddd|eee|fff|
				|ggg|hhh|iii|

				""";

			await Assert.That(result).IsEqualTo(expected);
		}

		private IEnumerator GetRangeEnumerator()
		{
			yield return selection.Object;
			yield return selection.Object;
			yield return selection.Object;
		}
	}
}
