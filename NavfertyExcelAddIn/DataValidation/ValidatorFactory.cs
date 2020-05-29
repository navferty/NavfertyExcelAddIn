using System;

using NavfertyExcelAddIn.DataValidation.Validators;

namespace NavfertyExcelAddIn.DataValidation
{
	public class ValidatorFactory : IValidatorFactory
	{
		public IValidator CreateValidator(ValidationType validationType)
		{
			switch (validationType)
			{
				case ValidationType.Numeric:
					{
						return new NumericValidator();
					}
				case ValidationType.Date:
					{
						return new DateValidator();
					}
				case ValidationType.TinPersonal:
					{
						return new TinPersonalValidator();
					}
				case ValidationType.TinOrganization:
					{
						return new TinOrganizationValidator();
					}
				case ValidationType.Xml:
					{
						return new XmlTextValidator();
					}
				default:
					{
						throw new ArgumentOutOfRangeException(nameof(validationType), validationType, null);
					}
			}
		}
	}
}
