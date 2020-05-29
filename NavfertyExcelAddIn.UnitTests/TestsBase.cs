using System.Globalization;
using System.IO;
using System.Threading;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using NavfertyExcelAddIn.Commons;

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

		protected string GetFilePath(string fileName)
		{
			return Path.Combine(Directory.GetCurrentDirectory(), fileName);
		}
	}
}
