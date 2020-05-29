namespace NavfertyExcelAddIn.Commons
{
	public class FileExtensionFilter
	{
		public FileExtensionFilter(string description, string extentions)
		{
			Description = description;
			Extentions = extentions;
		}

		public string Description { get; }
		public string Extentions { get; }
	}
}
