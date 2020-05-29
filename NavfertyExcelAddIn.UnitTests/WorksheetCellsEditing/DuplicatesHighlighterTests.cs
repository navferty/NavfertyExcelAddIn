using System.Linq;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using NavfertyExcelAddIn.UnitTests.Builders;
using NavfertyExcelAddIn.WorksheetCellsEditing;

namespace NavfertyExcelAddIn.UnitTests.WorksheetCellsEditing
{
	[TestClass]
	public class DuplicatesHighlighterTests : TestsBase
	{
		private DuplicatesHighlighter duplicatesHighlighter;

		[TestInitialize]
		public void BeforeEachTest()
		{
			SetRangeExtentionsStub();

			duplicatesHighlighter = new DuplicatesHighlighter();
		}

		[TestMethod]
		public void HighlightDuplicates_ThreeItems_Success()
		{
			var cells = new[]
			{
				new RangeBuilder().WithSingleValue("abc").Build(),
				new RangeBuilder().WithSingleValue("a").Build(),
				new RangeBuilder().WithSingleValue("b").Build(),
				new RangeBuilder().WithSingleValue("c").Build(),
				new RangeBuilder().WithSingleValue("a").Build(),
				new RangeBuilder().WithSingleValue("b").Build(),
				new RangeBuilder().WithSingleValue("qwe").Build(),
				new RangeBuilder().WithSingleValue(2d).Build(),
				new RangeBuilder().WithSingleValue(3d).Build(),
				new RangeBuilder().WithSingleValue(2d).Build()
			};
			var selection = new RangeBuilder().WithEnumrableRanges(cells).Build();


			duplicatesHighlighter.HighlightDuplicates(selection);

			Assert.AreEqual(3, RangeExtensionsStub.SetColorInvocations.Count());
		}

		[TestMethod]
		public void HighlightDuplicates_ManyItemsWithSameColors_Success()
		{
			var cells = Enumerable.Range(1, 1000).Select(x => new RangeBuilder().WithSingleValue((x % 100).ToString()).Build());
			var selection = new RangeBuilder().WithEnumrableRanges(cells).Build();

			duplicatesHighlighter.HighlightDuplicates(selection);

			// all 57 colors
			Assert.AreEqual(57, RangeExtensionsStub.SetColorInvocations.Count());
		}
	}
}
