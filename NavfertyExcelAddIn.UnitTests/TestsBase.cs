using System.IO;
using System.Collections;
using System.Globalization;
using System.Reflection;
using System.Threading;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.FindFormulaErrors;

using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace NavfertyExcelAddIn.UnitTests
{
    [TestClass]
    public abstract class TestsBase
    {
        protected RangeExtensionsImplementationStub RangeExtensionsStub { get; set; }

        public TestContext TestContext { get; set; }


        protected void SetRangeExtentionsStub()
        {
            RangeExtensionsStub = new RangeExtensionsImplementationStub();
            RangeExtensions.ResetImplementation(RangeExtensionsStub);
        }

        protected void SetCulture(string culture = "en-US")
        {
            var cultureInfo = CultureInfo.GetCultureInfo(culture);
            Thread.CurrentThread.CurrentCulture = cultureInfo;
            Thread.CurrentThread.CurrentUICulture = cultureInfo;
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

            var areas = new Mock<Areas>(MockBehavior.Strict);
            areas.As<IEnumerable>().Setup(x => x.GetEnumerator()).Returns(new[] { range.Object }.GetEnumerator());

            range.Setup(x => x.Areas).Returns(areas.Object);

            var cell = GetCellStub(values);

            range.Setup(x => x.Cells[It.IsAny<int>(), It.IsAny<int>()]).Returns(cell.Object);

            return range;
        }

        protected Mock<Range> GetCellStub(object value)
        {
            var cell = new Mock<Range>(MockBehavior.Loose);
            cell.Setup(x => x.get_Value(Missing.Value)).Returns(value);
            return cell;
        }

        protected string GetFilePath(string fileName)
        {
            return Path.Combine(Directory.GetCurrentDirectory(), fileName);
        }
    }
}
