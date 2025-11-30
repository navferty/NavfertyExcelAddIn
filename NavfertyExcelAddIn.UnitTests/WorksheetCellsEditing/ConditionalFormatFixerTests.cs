using System.Reflection;

using Microsoft.Office.Interop.Excel;

using Moq;

using NavfertyExcelAddIn.UnitTests.Builders;
using NavfertyExcelAddIn.WorksheetCellsEditing;

namespace NavfertyExcelAddIn.UnitTests.WorksheetCellsEditing;

public class ConditionalFormatFixerTests : TestsBase
{
	[Before(HookType.Test)]
	public void BeforeEachTest()
	{
		SetRangeExtentionsStub();
	}

	[Test]
	public void RepairConditionalFormat_OneRowSelected_NoAction()
	{
		var firstRow = new RangeBuilder();
		var rangeBuilder = new RangeBuilder()
			.WithConditionalFormatting(1)
			.WithWorksheet()
			.WithRows()
			.WithCount(1)
			.WithIndexer([firstRow.Build()])
			.WithPaste();
		var selection = rangeBuilder.Build();
		var formatFixer = new ConditionalFormatFixer();

		formatFixer.FillRange(selection);

		firstRow.MockObject.Verify(x => x.Copy(Missing.Value), Times.Never);
		rangeBuilder.MockObject.Verify(x => x.Copy(Missing.Value), Times.Never);
	}

	[Test]
	public void RepairConditionalFormat_TwoRowsSelected_FirstToSecond()
	{
		var firstRow = new RangeBuilder().WithConditionalFormatting(1).WithWorksheet().WithCopy();
		var secondRow = new RangeBuilder().WithConditionalFormatting(0).WithWorksheet().WithPaste();

		var rangeBuilder = new RangeBuilder()
			.WithRows()
			.WithCount(2)
			.WithIndexer([firstRow.Build(), secondRow.Build()]);
		var selection = rangeBuilder.Build();
		var formatFixer = new ConditionalFormatFixer();

		formatFixer.FillRange(selection);

		firstRow.MockObject.Verify(x => x.Copy(Missing.Value), Times.Once);
		VerifyPasted(secondRow.MockObject);
	}

	[Test]
	public void RepairConditionalFormat_TwoRowsSelected_SecondToFirst()
	{
		var firstRow = new RangeBuilder().WithConditionalFormatting(0).WithWorksheet().WithPaste();
		var secondRow = new RangeBuilder().WithConditionalFormatting(1).WithWorksheet().WithCopy();

		var rangeBuilder = new RangeBuilder()
			.WithRows()
			.WithCount(2)
			.WithIndexer([firstRow.Build(), secondRow.Build()]);
		var selection = rangeBuilder.Build();
		var formatFixer = new ConditionalFormatFixer();

		formatFixer.FillRange(selection);

		secondRow.MockObject.Verify(x => x.Copy(Missing.Value), Times.Once);
		VerifyPasted(firstRow.MockObject);
	}

	[Test]
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
			.WithIndexer([firstRow.Build(), secondRow.Build(), thirdRow.Build(), lastRow.Build()])
			.WithPaste();
		var selection = rangeBuilder.Build();
		var formatFixer = new ConditionalFormatFixer();

		formatFixer.FillRange(selection);

		firstRow.MockObject.Verify(x => x.Copy(Missing.Value), Times.Once);
		VerifyPasted(rangeBuilder.MockObject);
	}

	[Test]
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
			.WithIndexer([firstRow.Build(), secondRow.Build(), thirdRow.Build()])
			.WithCopy()
			.WithPaste();
		var selection = rangeBuilder.Build();
		var formatFixer = new ConditionalFormatFixer();

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
