using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.InteractiveRangeReport
{
    public class InteractiveErrorItem
    {
        public Range Range { get; set; }

        public string ErrorMessage { get; set; }
        public string Value { get; set; }
        public string Address { get; set; }
        public string WorksheetName { get; set; }
    }
}
