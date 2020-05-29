namespace NavfertyExcelAddIn.XmlTools
{
	public class XmlValidationError
	{
		public XmlValidationError(ValidationErrorSeverity severity, string message, string value, string elementName)
		{
			Severity = severity;
			Message = message;
			Value = value;
			ElementName = elementName;
		}

		public ValidationErrorSeverity Severity { get; set; }
		public string Message { get; set; }
		public string Value { get; set; }
		public string ElementName { get; set; }

		public override string ToString() => $"{Severity.ToString()}. {ElementName}: '{Value}': {Message}";
	}

	public enum ValidationErrorSeverity
	{
		Error,
		Warning
	}
}
