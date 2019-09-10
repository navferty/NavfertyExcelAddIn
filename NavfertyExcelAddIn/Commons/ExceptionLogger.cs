﻿using System;
using System.Diagnostics;
using NLog;
using Castle.DynamicProxy;
using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.Commons
{
    [DebuggerStepThrough]
    public class ExceptionLogger : IInterceptor
    {
        private readonly ILogger logger = LogManager.GetCurrentClassLogger();
        private readonly IDialogService dialogService;

        public ExceptionLogger(IDialogService dialogService)
        {
            this.dialogService = dialogService;
        }

        public void Intercept(IInvocation invocation)
        {
            try
            {
                invocation.Proceed();
            }
            catch (Exception ex)
            {
                logger.Error(ex);
                dialogService.ShowError(string.Format(UIStrings.ErrorMessage, ex.Message));
                throw;
            }
        }
    }
}
