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
		private TextTrimmer emptySpaceTrimmer;

		[TestInitialize]
		public void BeforeEachTest()
		{
			SetRangeExtentionsStub();

			emptySpaceTrimmer = new TextTrimmer();
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


			emptySpaceTrimmer.TrimExtraSpaces(selection);


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


			emptySpaceTrimmer.TrimExtraSpaces(selection);

			var expected = new object[,]
			{
				{ "abc", "def", "ghi" },
				{ "jk l", null, null },
				{ null, null, (int)CVErrEnum.ErrNA }
			};
			rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.IsAny<object[,]>()), Times.Once);
			rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.Is<object[,]>(v => AssertAssignedValue(expected, v))));
		}

		[TestMethod]
		public void RemoveAllSpaces_EmptyValue_Null()
		{
			var value = "   \r\n     \t";
			var rangeBuilder = new RangeBuilder()
				.WithWorksheet()
				.WithAreas()
				.WithSingleValue(value)
				.WithSetValue();
			var selection = rangeBuilder.Build();


			emptySpaceTrimmer.RemoveAllSpaces(selection);


			rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.IsAny<string>()), Times.Once);
			rangeBuilder.MockObject.Verify(x => x.set_Value(It.IsAny<object>(), It.Is<string>(v => v == null)));
		}

		[TestMethod]
		public void RemoveAllSpaces_AllValues_NoSpacesLeft()
		{
			var values = new object[,]
			{
				{ "abc ", " def", "  g h i\t" },
				{ "jk\r\nl", "   \r\n   \t", "   " },
				{ "", null, (int)CVErrEnum.ErrNA }
			};
			var rangeBuilder = new RangeBuilder()
				.WithWorksheet()
				.WithAreas()
				.WithMultipleValue(values)
				.WithSetValue();
			var selection = rangeBuilder.Build();


			emptySpaceTrimmer.RemoveAllSpaces(selection);

			var expected = new object[,]
			{
				{ "abc", "def", "ghi" },
				{ "jkl", null, null },
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
