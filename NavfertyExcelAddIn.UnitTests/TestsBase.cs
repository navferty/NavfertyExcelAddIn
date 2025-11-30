using System.Globalization;

using NavfertyExcelAddIn.Commons;

namespace NavfertyExcelAddIn.UnitTests;

public abstract class TestsBase
{
	protected RangeExtensionsStub RangeExtensionsStub { get; set; } = new();

	// TODO check when context may be null
	public TestContext TestContext { get; } = TestContext.Current!;


	protected void SetRangeExtentionsStub()
	{
		RangeExtensionsStub = new RangeExtensionsStub();
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
