using System.Reflection;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.FindFormulaErrors;

using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace NavfertyExcelAddIn.UnitTests
{
    [TestClass]
    public abstract class ExcelTests
    {
        protected Mock<IRangeExtensionsImplementation> rangeExtensions;

        [TestInitialize]
        public virtual void BeforeEachTest()
        {
            // TODO make separate stub class
            rangeExtensions = new Mock<IRangeExtensionsImplementation>(MockBehavior.Strict);
            rangeExtensions.Setup(x => x.GetFormula(It.IsAny<Range>())).Returns("=1+1");
            rangeExtensions.Setup(x => x.GetRelativeAddress(It.IsAny<Range>())).Returns("A1");
            rangeExtensions.Setup(x => x.SetColor(It.IsAny<Range>(), It.IsAny<int>()));

            RangeExtensions.ResetImplementation(rangeExtensions.Object);
        }

        // TODO make as builder 'Builder.WithValue(value)' etc.
        protected Mock<Range> GetRangeStub(object[,] values = null)
        {
            var ws = new Mock<Worksheet>(MockBehavior.Loose);
            ws.Setup(x => x.Name).Returns("Sheet1");

            var range = new Mock<Range>(MockBehavior.Loose);

            range.Setup(x => x.Worksheet)
                .Returns(ws.Object);

            range.Setup(x => x.get_Value(Missing.Value))
                .Returns(values ?? new object[,] { { "asd", "dsa" }, { 123.456d, (int)CVErrEnum.ErrNA } });

            var cell = GetCellStub(values);

            range.Setup(x => x.Cells[It.IsAny<int>(), It.IsAny<int>()]).Returns(cell.Object);

            return range;
        }

        protected Mock<Range> GetCellStub(object value = null)
        {
            var cell = new Mock<Range>(MockBehavior.Loose);
            cell.Setup(x => x.get_Value(Missing.Value)).Returns(value ?? "asd");
            return cell;
        }
    }
}
