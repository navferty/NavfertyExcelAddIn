﻿using System.ComponentModel;

namespace NavfertyExcelAddIn.FindFormulaErrors
{
    public enum CVErrEnum : int
    {
        
        [Description("#DIV/0!")]
        ErrDiv0 = -2146826281,
        
        [Description("#GETTING_DATA")]
        ErrGettingData = -2146826245,

        [Description("#N/A")]
        ErrNA = -2146826246,

        [Description("#NAME?")]
        ErrName = -2146826259,

        [Description("#NULL!")]
        ErrNull = -2146826288,

        [Description("#NUM!")]
        ErrNum = -2146826252,

        [Description("#REF!")]
        ErrRef = -2146826265,

        [Description("#VALUE!")]
        ErrValue = -2146826273
    }
}
