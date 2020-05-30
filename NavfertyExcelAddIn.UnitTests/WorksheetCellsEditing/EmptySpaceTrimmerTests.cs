using Microsoft.VisualStudio.TestTools.UnitTesting;

using Moq;

using NavfertyExcelAddIn.FindFormulaErrors;
using NavfertyExcelAddIn.UnitTests.Builders;
using NavfertyExcelAddIn.WorksheetCellsEditing;

namespace NavfertyExcelAddIn.UnitTests.WorksheetCellsEditing
{
	[TestClass]
	public class EmptySpaceTrimmerTests : TestsBase
	{
		private EmptySpaceTrimmer emptySpaceTrimmer;

		[TestInitialize]
		public void BeforeEachTest()
		{
			SetRangeExtentionsStub();

			emptySpaceTrimmer = new EmptySpaceTrimmer();
		}

		[TestMethod]
		public void TrimSpaces_SingleCell_Converted()
		{
			var value = "a   \r\nb     c\t";
			var rangeBuilder = new RangeBuilder()
				.WithWorksheet()
				.WithAreas()
				.WithSingleValue(value)
				.WithSetValue();
			var selection = rangeBuilder.Build();


			emptySpaceTrimmer.TrimSpaces(selection);


			var expected = "a b c";
			rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.IsAny<string>()), Times.Once);
			rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.Is<string>(v => expected == v)));
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
			var rangeBuilder = new RangeBuilder()
				.WithWorksheet()
				.WithAreas()
				.WithMultipleValue(values)
				.WithSetValue();
			var selection = rangeBuilder.Build();


			emptySpaceTrimmer.TrimSpaces(selection);

			var expected = new object[,]
			{
				{ "abc", "def", "ghi" },
				{ "jk l", null, null },
				{ null, null, (int)CVErrEnum.ErrNA }
			};
			rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.IsAny<object[,]>()), Times.Once);
			rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.Is<object[,]>(v => AssertAssignedValue(expected, v))));
		}
		private bool AssertAssignedValue(object[,] expected, object[,] value)
		{
			CollectionAssert.AreEquivalent(expected, value);
			return true;
		}
	}
}
