namespace NavfertyExcelAddIn.DataValidation
{
    public interface IValidatorFactory
    {
        IValidator CreateValidator(ValidationType validationType);
    }
}