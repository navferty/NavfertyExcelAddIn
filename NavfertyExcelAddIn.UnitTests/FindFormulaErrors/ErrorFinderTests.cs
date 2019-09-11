using System;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

using NavfertyExcelAddIn.FindFormulaErrors;
using NavfertyExcelAddIn.Commons;

using Microsoft.Office.Interop.Excel;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Range = Microsoft.Office.Interop.Excel.Range;


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
        [Ignore("Screw this shit with dynamics")]
        public void Range_WithErrors()
        {
            errorFinder = new ErrorFinder();

            Mock<Range> worksheetUsedRange = SetUpRangeWithValue(
                new object[,] { { 1d, "1", "abc" }, { (int)CVErrEnum.ErrGettingData, "321 , 321", null } });
            worksheetUsedRange.Setup(r => r.Cells[It.IsAny<int>(), It.IsAny<int>()]).Returns(worksheetUsedRange.Object);

            var result = errorFinder.GetAllErrorCells(worksheetUsedRange.Object).ToArray();

            Assert.AreEqual(1, result.Length);
            Assert.AreEqual(CVErrEnum.ErrGettingData, result.First().ErrorMessage);
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
        [Ignore("Screw this shit with dynamics")]
        public void SingleValue_WithError()
        {
            errorFinder = new ErrorFinder();

            Mock<Range> worksheetUsedRange = SetUpRangeWithValue((int)CVErrEnum.ErrNA);

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

        private static Mock<Range> SetUpRangeWithValue(object rangeValue)
        {
            var ws = new Mock<Worksheet>(MockBehavior.Strict);
            ws.Setup(x => x.Name).Returns("WsName");
            var worksheetUsedRange = new Mock<Range>(MockBehavior.Loose);
            worksheetUsedRange.SetupGet(x => x.get_Value(Missing.Value)).Returns(rangeValue);

            Expression<Func<Range, string>> getAddress = x => x.get_Address(It.IsAny<object>(), It.IsAny<object>(),
                                It.IsAny<XlReferenceStyle>(), It.IsAny<object>(), It.IsAny<object>());

            worksheetUsedRange.SetupGet(getAddress).Returns("Address");

            //var getFormula = GetFormulaExpression();
            //worksheetUsedRange.Setup(getFormula).Returns("=SomeFormula");

            worksheetUsedRange.Setup(x => x.Worksheet).Returns(ws.Object);
            return worksheetUsedRange;
        }

        //private static Expression<Func<Range, object>> GetFormulaExpression()
        //{
        //    ParameterExpression target = Expression.Parameter(typeof(Range), "target");
        //    ParameterExpression result = Expression.Parameter(typeof(object), "result");
        //    CallSiteBinder getFormula = Binder.GetMember(
        //        CSharpBinderFlags.None,
        //        "Formula",
        //        typeof(ThisAddIn),
        //        new CSharpArgumentInfo[] { CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null) });
        //    var expression = Expression.Lambda<Func<Range, object>>(Expression.Dynamic(getFormula, typeof(object), target), new[] { target });
        //    return expression;
        //}
    }
}
