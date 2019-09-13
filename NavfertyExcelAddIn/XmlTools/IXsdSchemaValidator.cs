using System.Collections.Generic;

namespace NavfertyExcelAddIn.XmlTools
{
    public interface IXsdSchemaValidator
    {
        IReadOnlyCollection<string> Validate(string xmlFilename, IReadOnlyCollection<string> schemas);
    }
}
