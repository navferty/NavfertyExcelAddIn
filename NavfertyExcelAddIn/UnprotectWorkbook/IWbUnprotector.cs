using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NavfertyExcelAddIn.UnprotectWorkbook
{
    public interface IWbUnprotector
    {
        void UnprotectWorkbookWithAllWorksheets(string path);
    }
}
