using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;

using Application = Microsoft.Office.Interop.Excel.Application;

namespace NavfertyExcelAddIn
{
    [ComVisible(true)]
    [SuppressMessage("Стиль", "IDE0060:Удалите неиспользуемый параметр", Justification = "Method signatures are public")]
    public class NavfertyRibbon : IRibbonExtensibility
    {
        private Application App => Globals.ThisAddIn.Application;

        public NavfertyRibbon()
        {
            // TODO
        }

        #region IRibbonExtensibility
        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText();
        }
        #endregion

        #region Ribbon callbacks
        public void RibbonLoad(IRibbonUI ribbonUI)
        {
            // TODO
        }

        public void ParseNumerics(IRibbonControl ribbonControl)
        {
            var wb = App.ActiveWorkbook;
            MessageBox.Show($"Parse numerics in workbook {wb.Name}. Not implemented yet");
        }

        public Bitmap GetImage(string imageName)
        {
            return (Bitmap)RibbonImages.ResourceManager.GetObject(imageName);
        }
        #endregion

        #region Utils
        private static string GetResourceText()
        {
            var asm = typeof(NavfertyRibbon).Assembly;

            using (var stream = asm.GetManifestResourceStream("NavfertyExcelAddIn.NavfertyRibbon.xml"))
            using (var resourceReader = new StreamReader(stream))
            {
                return resourceReader.ReadToEnd();
            }
        }
        #endregion
    }
}
