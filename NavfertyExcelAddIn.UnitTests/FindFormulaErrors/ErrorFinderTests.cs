using Moq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Range = Microsoft.Office.Interop.Excel.Range;
using NavfertyExcelAddIn.FindFormulaErrors;
using System.Linq;
using System.Reflection;

namespace NavfertyExcelAddIn.UnitTests.FindFormulaErrors
{
    [TestClass]
    public class ErrorFinderTests
    {
        private ErrorFinder errorFinder;

        [TestMethod]
        public void Range_NoErrors()
        {
            errorFinder = new ErrorFinder();

            Mock<Range> worksheetUsedRange = SetUpRangeWithValue(
                new object[,] { { 1d, "1", "abc" }, { "123.123", "321 , 321", null } });

            var result = errorFinder.GetAllErrorCells(worksheetUsedRange.Object).ToArray();

            Assert.AreEqual(0, result.Length);
        }

        [TestMethod]
        public void Range_WithErrors()
        {
            errorFinder = new ErrorFinder();

            Mock<Range> worksheetUsedRange = SetUpRangeWithValue(
                new object[,] { { 1d, "1", "abc" }, { (int)CVErrEnum.ErrGettingData, "321 , 321", null } });
            worksheetUsedRange.Setup(r => r.Cells[It.IsAny<int>(), It.IsAny<int>()]).Returns(worksheetUsedRange.Object);

            var result = errorFinder.GetAllErrorCells(worksheetUsedRange.Object).ToArray();

            Assert.AreEqual(1, result.Length);
            Assert.AreEqual(CVErrEnum.ErrGettingData, result.First().ErrorType);
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

            var result = errorFinder.GetAllErrorCells(worksheetUsedRange.Object).ToArray();

            Assert.AreEqual(1, result.Length);
            Assert.AreEqual(CVErrEnum.ErrNA, result.First().ErrorType);
        }

        [TestMethod]
        public void NoValues()
        {
            errorFinder = new ErrorFinder();

            Mock<Range> worksheetUsedRange = SetUpRangeWithValue(null);

            var result = errorFinder.GetAllErrorCells(worksheetUsedRange.Object).ToArray();

            Assert.AreEqual(0, result.Length);
        }

        private static Mock<Range> SetUpRangeWithValue(object rangeValue)
        {
            var worksheetUsedRange = new Mock<Range>(MockBehavior.Strict);
            worksheetUsedRange.SetupGet(x => x.get_Value(Missing.Value)).Returns(rangeValue);
            return worksheetUsedRange;
        }
    }
}
