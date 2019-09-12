namespace NavfertyExcelAddIn.DataValidation
{
    public interface IValidator
    {
		ValidationResult CheckValue(object cellValue);
    }
}
