using System.Collections;
using System.Reflection;

using Moq;

using NavfertyExcelAddIn.FindFormulaErrors;

using Areas = Microsoft.Office.Interop.Excel.Areas;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;

namespace NavfertyExcelAddIn.UnitTests.Builders;

public class RangeBuilder : TestDataBuilder<Range>
{
	private Range?[]? ranges;

	public Mock<Range> MockObject { get; private set; } = new Mock<Range>(MockBehavior.Strict);

	public RangeBuilder WithWorksheet(string wsName = "Sheet1")
	{
		var ws = new Mock<Worksheet>(MockBehavior.Loose);
		ws.Setup(x => x.Name).Returns(wsName);

		var app = new Mock<Application>(MockBehavior.Strict);
		app.SetupGet(x => x.CutCopyMode).Returns(Microsoft.Office.Interop.Excel.XlCutCopyMode.xlCopy);
		app.SetupSet(x => x.CutCopyMode = It.IsAny<Microsoft.Office.Interop.Excel.XlCutCopyMode>());

		ws.Setup(x => x.Application).Returns(app.Object);

		ws.SetupGet(x => x.get_Range(It.IsAny<object>(), It.IsAny<object>()))
			.Returns(MockObject.Object);

		var parentWb = new Mock<Workbook>(MockBehavior.Loose);
		parentWb.Setup(x => x.Name).Returns("Book1");
		ws.Setup(x => x.Parent).Returns(parentWb);
		MockObject.Setup(x => x.Worksheet).Returns(ws.Object);

		return this;
	}

	public RangeBuilder WithCells()
	{
		MockObject.Setup(x => x.Cells[It.IsAny<int>(), It.IsAny<int>()]).Returns(MockObject.Object);
		return this;
	}

	public RangeBuilder WithAreas()
	{
		var areas = new Mock<Areas>(MockBehavior.Strict);
		areas.Setup(x => x.GetEnumerator())
			.Returns(new[] { MockObject.Object }.GetEnumerator());

		MockObject.Setup(x => x.Areas).Returns(areas.Object);
		MockObject.Setup(x => x.Row).Returns(42);
		MockObject.Setup(x => x.Column).Returns(42);
		return this;
	}

	public RangeBuilder WithIndexer(IEnumerable<Range?>? ranges = null)
	{
		this.ranges = ranges?.ToArray()
			?? Enumerable.Range(0, 10).Select(z => new RangeBuilder().WithSingleValue(z).Build()).ToArray();
#pragma warning disable CS8603 // Possible null reference return.
		MockObject
			.Setup(x => x[It.IsAny<int>(), Missing.Value])
			.Returns(new InvocationFunc(i => this.ranges[(int)i.Arguments.First() - 1]));
#pragma warning restore CS8603 // Possible null reference return.
		return this;
	}

	public RangeBuilder WithRows()
	{
		MockObject.Setup(x => x.Rows).Returns(MockObject.Object);
		return this;
	}

	public RangeBuilder WithCopy()
	{
		MockObject.Setup(x => x.Copy(It.IsAny<Missing>())).Returns(MockObject.Object);
		return this;
	}

	public RangeBuilder WithPaste()
	{
		MockObject
			.Setup(x => x.PasteSpecial(
				It.IsAny<Microsoft.Office.Interop.Excel.XlPasteType>(),
				It.IsAny<Microsoft.Office.Interop.Excel.XlPasteSpecialOperation>(),
				It.IsAny<object>(),
				It.IsAny<object>()))
			.Returns(MockObject.Object);

		return this;
	}

	public RangeBuilder WithCount(int count)
	{
		MockObject.Setup(x => x.Count).Returns(count);
		return this;
	}

	public RangeBuilder WithConditionalFormatting(int rulesCount)
	{
		var f = new Mock<Microsoft.Office.Interop.Excel.FormatConditions>(MockBehavior.Strict);
		f.Setup(x => x.Count).Returns(rulesCount);
		f.Setup(x => x.Delete());
		MockObject.Setup(x => x.FormatConditions).Returns(f.Object);
		return this;
	}

	public RangeBuilder WithSingleValue(object? value)
	{
		MockObject.Setup(x => x.get_Value(Missing.Value)).Returns(value!);
		return this;
	}

	public RangeBuilder WithMultipleValue(object?[,] values)
	{
		var value = values ?? new object[,] { { "asd", "dsa" }, { 123.456d, (int)CVErrEnum.ErrNA } };
		MockObject.Setup(x => x.get_Value(Missing.Value)).Returns(value);
		return this;
	}

	public RangeBuilder WithSetValue()
	{
		MockObject.Setup(x => x.set_Value(It.IsAny<object>(), It.IsAny<object>()));
		return this;
	}

	public RangeBuilder WithEnumrableRanges(IEnumerable<Range?>? ranges = null)
	{
		this.ranges = ranges?.ToArray()
			?? Enumerable.Range(0, 10).Select(z => new RangeBuilder().WithSingleValue(z).Build()).ToArray();

		MockObject
			.As<IEnumerable>()
			.Setup(x => x.GetEnumerator())
			.Returns(this.ranges.GetEnumerator());
		return this;
	}

	public override Range Build()
	{
		return MockObject.Object;
	}
}
