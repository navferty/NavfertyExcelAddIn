using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Moq;

using NavfertyExcelAddIn.FindFormulaErrors;

using Areas = Microsoft.Office.Interop.Excel.Areas;
using Range = Microsoft.Office.Interop.Excel.Range;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace NavfertyExcelAddIn.UnitTests.Builders
{
    public class RangeBuilder : TestDataBuilder<Range>
    {
        private Mock<Worksheet> ws;
        private Mock<Areas> areas;

        public Mock<Range> MockObject { get; private set; } = new Mock<Range>(MockBehavior.Strict);

        public RangeBuilder WithWorksheet(string wsName = "Sheet1")
        {
            ws = new Mock<Worksheet>(MockBehavior.Loose);
            ws.Setup(x => x.Name).Returns(wsName);
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
            areas = new Mock<Areas>(MockBehavior.Strict);
            areas.As<IEnumerable>().Setup(x => x.GetEnumerator())
                .Returns(new[] { MockObject.Object }.GetEnumerator());

            MockObject.Setup(x => x.Areas).Returns(areas.Object);
            return this;
        }

        public RangeBuilder WithSingleValue(object value)
        {
            MockObject.Setup(x => x.get_Value(Missing.Value)).Returns(value);
            return this;
        }

        public RangeBuilder WithMultipleValue(object[,] values)
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

        public RangeBuilder WithEnumrableRanges(IEnumerable<Range> ranges = null)
        {
            var r = ranges ?? Enumerable.Range(0, 10).Select(z => new RangeBuilder().WithSingleValue(z).Build());

            MockObject
                .As<IEnumerable>()
                .Setup(x => x.GetEnumerator())
                .Returns(r.GetEnumerator());
            return this;
        }

        public override Range Build()
        {
            return MockObject.Object;
        }
    }
}
