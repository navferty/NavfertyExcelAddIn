using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using Moq;

using NavfertyExcelAddIn.UnitTests.Builders;
using NavfertyExcelAddIn.WorksheetCellsEditing;

namespace NavfertyExcelAddIn.UnitTests.WorksheetCellsEditing
{
	[TestClass]
	public class ConditionalFormatFixerTests : TestsBase
	{
		private ConditionalFormatFixer formatFixer;

		[TestInitialize]
		public void BeforeEachTest()
		{
			SetRangeExtentionsStub();

			formatFixer = new ConditionalFormatFixer();
		}

		[TestMethod]
		public void RepairConditionalFormat_OneRowSelected_NoAction()
		{
			var firstRow = new RangeBuilder();
			var rangeBuilder = new RangeBuilder()
				.WithConditionalFormatting(1)
				.WithWorksheet()
				.WithRows()
				.WithCount(1)
				.WithIndexer(new[] { firstRow.Build() })
				.WithPaste();
			var selection = rangeBuilder.Build();

			formatFixer.FillRange(selection);

			firstRow.MockObject.Verify(x => x.Copy(Missing.Value), Times.Never);
			rangeBuilder.MockObject.Verify(x => x.Copy(Missing.Value), Times.Never);
		}

		[TestMethod]
		public void RepairConditionalFormat_TwoRowsSelected_FirstToSecond()
		{
			var firstRow = new RangeBuilder().WithConditionalFormatting(1).WithWorksheet().WithCopy();
			var secondRow = new RangeBuilder().WithConditionalFormatting(0).WithWorksheet().WithPaste();

			var rangeBuilder = new RangeBuilder()
				.WithRows()
				.WithCount(2)
				.WithIndexer(new[] { firstRow.Build(), secondRow.Build() });
			var selection = rangeBuilder.Build();

			formatFixer.FillRange(selection);

			firstRow.MockObject.Verify(x => x.Copy(Missing.Value), Times.Once);
			VerifyPasted(secondRow.MockObject);
		}

		[TestMethod]
		public void RepairConditionalFormat_TwoRowsSelected_SecondToFirst()
		{
			var firstRow = new RangeBuilder().WithConditionalFormatting(0).WithWorksheet().WithPaste();
			var secondRow = new RangeBuilder().WithConditionalFormatting(1).WithWorksheet().WithCopy();

			var rangeBuilder = new RangeBuilder()
				.WithRows()
				.WithCount(2)
				.WithIndexer(new[] { firstRow.Build(), secondRow.Build() });
			var selection = rangeBuilder.Build();

			formatFixer.FillRange(selection);

			secondRow.MockObject.Verify(x => x.Copy(Missing.Value), Times.Once);
			VerifyPasted(firstRow.MockObject);
		}

		[TestMethod]
		public void RepairConditionalFormat_MultipleRows_CopiedFromFirstRow()
		{
			var firstRow = new RangeBuilder().WithConditionalFormatting(1).WithWorksheet().WithCopy();
			var secondRow = new RangeBuilder().WithConditionalFormatting(0).WithWorksheet();
			var thirdRow = new RangeBuilder().WithConditionalFormatting(1).WithWorksheet();
			var lastRow = new RangeBuilder().WithConditionalFormatting(0).WithWorksheet();

			var rangeBuilder = new RangeBuilder()
				.WithConditionalFormatting(1)
				.WithWorksheet()
				.WithRows()
				.WithCount(4)
				.WithIndexer(new[] { firstRow.Build(), secondRow.Build(), thirdRow.Build(), lastRow.Build() })
				.WithPaste();
			var selection = rangeBuilder.Build();

			formatFixer.FillRange(selection);

			firstRow.MockObject.Verify(x => x.Copy(Missing.Value), Times.Once);
			VerifyPasted(rangeBuilder.MockObject);
		}

		[TestMethod]
		public void RepairConditionalFormat_FirstRowWithoutFormat_CopiedFromSecondRow()
		{
			var firstRow = new RangeBuilder().WithConditionalFormatting(0).WithWorksheet();
			var secondRow = new RangeBuilder().WithConditionalFormatting(2).WithWorksheet().WithCopy();
			var thirdRow = new RangeBuilder().WithConditionalFormatting(0).WithWorksheet();

			var rangeBuilder = new RangeBuilder()
				.WithConditionalFormatting(1)
				.WithWorksheet()
				.WithRows()
				.WithCount(3)
				.WithIndexer(new[] { firstRow.Build(), secondRow.Build(), thirdRow.Build() })
				.WithCopy()
				.WithPaste();
			var selection = rangeBuilder.Build();

			formatFixer.FillRange(selection);

			secondRow.MockObject.Verify(x => x.Copy(Missing.Value), Times.Once);
			VerifyPasted(rangeBuilder.MockObject);
		}

		private static void VerifyPasted(Mock<Microsoft.Office.Interop.Excel.Range> range)
		{
			range.Verify(x => x.PasteSpecial(
				XlPasteType.xlPasteFormats,
				It.IsAny<XlPasteSpecialOperation>(),
				Missing.Value,
				Missing.Value), Times.Once);
		}
	}
}
