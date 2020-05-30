namespace NavfertyExcelAddIn.UnitTests.Builders
{
	public abstract class TestDataBuilder<TItem>
		where TItem : class
	{
		public abstract TItem Build();
	}
}
