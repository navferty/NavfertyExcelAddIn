using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.XmlTools
{
    public interface IXmlValidator
    {
        void Validate(Application excelApp);
    }
}
