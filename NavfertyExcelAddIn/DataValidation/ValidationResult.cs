namespace NavfertyExcelAddIn.DataValidation
{
	public class ValidationResult
	{
		public bool IsSuccess { get; set; }
		public string Message { get; set; } = string.Empty;

		public static ValidationResult Success => new ValidationResult { IsSuccess = true };
		public static ValidationResult Fail(string message) => new ValidationResult { IsSuccess = false, Message = message };
	}
}
