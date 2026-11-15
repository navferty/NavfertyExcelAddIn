using Moq;

using NavfertyExcelAddIn.Undo;
using NavfertyExcelAddIn.UnitTests.Builders;

namespace NavfertyExcelAddIn.UnitTests.Commons;

public class UndoManagerTests
{
	[Test]
	public void Undo_SingleCell_RestoresFormula()
	{
		var builder = new RangeBuilder().WithWorksheet().WithCells().WithSingleValue("three");
		var cell = builder.Build();
		var sheet = Mock.Get(cell.Worksheet);
		sheet.SetupGet(x => x.Cells).Returns(cell);
		builder.MockObject.SetupSet(x => x.Formula = It.IsAny<object>());
		var manager = new UndoManager();
		manager.StartNewAction(cell);
		manager.PushUndoItem(new UndoItem
		{
			RowIndex = 1,
			ColumnIndex = 1,
			OldFormula = "=A1+A2",
			NewValue = "three"
		});

		manager.UndoLastAction(sheet.Object);

		builder.MockObject.VerifySet(x => x.Formula = "=A1+A2", Times.Once);
	}

	[Test]
	public void Undo_Area_RestoresFormulasAndValues()
	{
		var newValues = new object[,] { { "one", "two" }, { "three", "text" } };
		var oldFormulas = new object[,] { { 1d, 2d }, { "=A1+B1", "text" } };
		var builder = new RangeBuilder().WithWorksheet().WithCells().WithMultipleValue(newValues);
		var range = builder.Build();
		var sheet = Mock.Get(range.Worksheet);
		sheet.SetupGet(x => x.Cells).Returns(range);
		builder.MockObject.SetupSet(x => x.Formula = It.IsAny<object>());
		var manager = new UndoManager();
		manager.StartNewAction(range);
		manager.PushAreaUndoItem(new AreaUndoItems
		{
			Row = 1,
			Column = 1,
			Height = 2,
			Width = 2,
			OldFormulas = oldFormulas,
			NewValues = newValues
		});

		manager.UndoLastAction(sheet.Object);

		builder.MockObject.VerifySet(x => x.Formula = oldFormulas, Times.Once);
	}

	[Test]
	public void Undo_ChangedCell_DoesNotRestoreFormula()
	{
		var builder = new RangeBuilder().WithWorksheet().WithCells().WithSingleValue("edited");
		var cell = builder.Build();
		var sheet = Mock.Get(cell.Worksheet);
		sheet.SetupGet(x => x.Cells).Returns(cell);
		var manager = new UndoManager();
		manager.StartNewAction(cell);
		manager.PushUndoItem(new UndoItem
		{
			RowIndex = 1,
			ColumnIndex = 1,
			OldFormula = "=A1+A2",
			NewValue = "three"
		});

		manager.UndoLastAction(sheet.Object);

		builder.MockObject.VerifySet(x => x.Formula = It.IsAny<object>(), Times.Never);
	}
}
