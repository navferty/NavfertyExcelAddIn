using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.DataValidation
{
    public class ValidationError
    {
        public Range Range { get; set; }
        public string ErrorMessage { get; set; }
        public string Info { get; set; }
        public string WorksheetName { get; set; }
    }
}