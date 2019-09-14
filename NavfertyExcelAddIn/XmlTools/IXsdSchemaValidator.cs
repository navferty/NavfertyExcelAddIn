using System.Collections.Generic;

namespace NavfertyExcelAddIn.XmlTools
{
    public interface IXsdSchemaValidator
    {
        IReadOnlyCollection<XmlValidationError> Validate(string xmlFilename, IReadOnlyCollection<string> schemas);
    }
}
