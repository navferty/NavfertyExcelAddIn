using System.Linq;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.FindFormulaErrors;
using NavfertyExcelAddIn.UnitTests.Builders;

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
			var range = new RangeBuilder().WithWorksheet().WithSingleValue(values).Build();

			var result = errorFinder.GetAllErrorCells(range).ToArray();

			Assert.AreEqual(0, result.Length);
		}

		[TestMethod]
		public void Range_WithErrors()
		{
			errorFinder = new ErrorFinder();

			var values = new object[,] { { 1d, "1", "abc" }, { (int)CVErrEnum.ErrGettingData, "321 , 321", null } };
			var range = new RangeBuilder().WithWorksheet().WithCells().WithSingleValue(values).Build();
			SetRangeExtentionsStub();

			var result = errorFinder.GetAllErrorCells(range).ToArray();

			Assert.AreEqual(1, result.Length);
			Assert.AreEqual(CVErrEnum.ErrGettingData.GetEnumDescription(), result.First().ErrorMessage);
		}

		[TestMethod]
		public void SingleValue_WithoutError()
		{
			errorFinder = new ErrorFinder();

			var range = new RangeBuilder().WithWorksheet().WithSingleValue("1").Build();

			var result = errorFinder.GetAllErrorCells(range).ToArray();

			Assert.AreEqual(0, result.Length);
		}

		[TestMethod]
		public void SingleValue_WithError()
		{
			errorFinder = new ErrorFinder();

			var range = new RangeBuilder().WithWorksheet().WithSingleValue((int)CVErrEnum.ErrNA).Build();

			SetRangeExtentionsStub();


			var result = errorFinder.GetAllErrorCells(range).ToArray();

			Assert.AreEqual(1, result.Length);
			Assert.AreEqual(CVErrEnum.ErrNA.GetEnumDescription(), result.First().ErrorMessage);
		}

		[TestMethod]
		public void NoValues()
		{
			errorFinder = new ErrorFinder();

			var range = new RangeBuilder().WithWorksheet().WithSingleValue(null).Build();

			var result = errorFinder.GetAllErrorCells(range).ToArray();

			Assert.AreEqual(0, result.Length);
		}
	}
}
