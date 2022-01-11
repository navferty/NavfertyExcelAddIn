using System;
using System.Collections.Generic;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.WorksheetProtectorUnprotector
{
    internal class WorksheetRow
    {
        public readonly Worksheet Sheet;

        public WorksheetRow(Worksheet ws)
        {
            Sheet = ws;
        }

        [Flags]
        public enum WorksheetProtectionFlags : int
        {
            None = 0,
            DrawingObjects = 1,
            Contents = 2,
            Scenarios = 4,
            UI = 8
        }

        public WorksheetProtectionFlags GetProtectionFlags()
        {
            WorksheetProtectionFlags f = WorksheetProtectionFlags.None;
            if (Sheet.ProtectDrawingObjects) f |= WorksheetProtectionFlags.DrawingObjects;
            if (Sheet.ProtectContents) f |= WorksheetProtectionFlags.Contents;
            if (Sheet.ProtectScenarios) f |= WorksheetProtectionFlags.Scenarios;
            if (Sheet.ProtectionMode) f |= WorksheetProtectionFlags.UI;
            return f;
        }

        public string GetProtectionFlagsLocalized()
        {
            WorksheetProtectionFlags f = GetProtectionFlags();
            var lpf = new List<string>();
            if (f.HasFlag(WorksheetProtectionFlags.DrawingObjects)) lpf.Add(UIStrings.SheetProtectionFlag_DrawingObjects);
            if (f.HasFlag(WorksheetProtectionFlags.Contents)) lpf.Add(UIStrings.SheetProtectionFlag_Contents);
            if (f.HasFlag(WorksheetProtectionFlags.Scenarios)) lpf.Add(UIStrings.SheetProtectionFlag_Scenarios);
            if (f.HasFlag(WorksheetProtectionFlags.UI)) lpf.Add(UIStrings.SheetProtectionFlag_UI);

            var result = string.Join(", ", lpf.ToArray());
            return result;
        }

        public bool HasAnyProtectedObjects() => GetProtectionFlags() != WorksheetProtectionFlags.None;


        public override string ToString()
        {
            var sProtection = HasAnyProtectedObjects() ? $" ({UIStrings.SheetProtection_Protected.ToLower().Trim()}: {GetProtectionFlagsLocalized()})" : "";
            return $"{Sheet.Name}{sProtection}";
        }
    }
}
