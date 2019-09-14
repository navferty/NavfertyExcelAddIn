using System.Linq;
using System.Reflection;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

using NavfertyExcelAddIn.FindFormulaErrors;
using NavfertyExcelAddIn.Commons;

using Range = Microsoft.Office.Interop.Excel.Range;

namespace NavfertyExcelAddIn.UnitTests.FindFormulaErrors
{
    [TestClass]
    public class ErrorFinderTests : TestsBase
    {
        private ErrorFinder errorFinder;

        [TestMethod]
        public void Range_NoErrors()
        {
            errorFinder = new ErrorFinder();

            var values = new object[,] { { 1d, "1", "abc" }, { "123.123", "321 , 321", null } };
            var worksheetUsedRange = SetUpRangeWithValue(values);

            var result = errorFinder.GetAllErrorCells(worksheetUsedRange.Object).ToArray();

            Assert.AreEqual(0, result.Length);
        }

        [TestMethod]
        public void Range_WithErrors()
        {
            errorFinder = new ErrorFinder();

            var values = new object[,] { { 1d, "1", "abc" }, { (int)CVErrEnum.ErrGettingData, "321 , 321", null } };
            var worksheetUsedRange = SetUpRangeWithValue(values);
            SetRangeExtentionsStub();

            var result = errorFinder.GetAllErrorCells(worksheetUsedRange.Object).ToArray();

            Assert.AreEqual(1, result.Length);
            Assert.AreEqual(CVErrEnum.ErrGettingData.GetEnumDescription(), result.First().ErrorMessage);
        }

        [TestMethod]
        public void SingleValue_WithoutError()
        {
            errorFinder = new ErrorFinder();

            Mock<Range> worksheetUsedRange = SetUpRangeWithValue("1");

            var result = errorFinder.GetAllErrorCells(worksheetUsedRange.Object).ToArray();

            Assert.AreEqual(0, result.Length);
        }

        [TestMethod]
        public void SingleValue_WithError()
        {
            errorFinder = new ErrorFinder();

            Mock<Range> worksheetUsedRange = SetUpRangeWithValue((int)CVErrEnum.ErrNA);

            SetRangeExtentionsStub();


            var result = errorFinder.GetAllErrorCells(worksheetUsedRange.Object).ToArray();

            Assert.AreEqual(1, result.Length);
            Assert.AreEqual(CVErrEnum.ErrNA.GetEnumDescription(), result.First().ErrorMessage);
        }

        [TestMethod]
        public void NoValues()
        {
            errorFinder = new ErrorFinder();

            Mock<Range> worksheetUsedRange = SetUpRangeWithValue(null);

            var result = errorFinder.GetAllErrorCells(worksheetUsedRange.Object).ToArray();

            Assert.AreEqual(0, result.Length);
        }

        private Mock<Range> SetUpRangeWithValue(object rangeValue)
        {
            var range = GetRangeStub();
            range.SetupGet(x => x.get_Value(Missing.Value)).Returns(rangeValue);
            return range;
        }
    }
}
