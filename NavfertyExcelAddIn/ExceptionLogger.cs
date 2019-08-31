﻿using System;
using NLog;
using Castle.DynamicProxy;
using System.Windows.Forms;
using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn
{
    public class ExceptionLogger : IInterceptor
    {
        private readonly ILogger logger = LogManager.GetCurrentClassLogger();

        public void Intercept(IInvocation invocation)
        {
            try
            {
                invocation.Proceed();
            }
            catch (Exception ex)
            {
                logger.Error(ex);
                MessageBox.Show(string.Format(UIStrings.ErrorMessage, ex.Message));
                throw;
            }
        }
    }
}
