using Autofac;
using Autofac.Extras.DynamicProxy;

using NavfertyExcelAddIn.ParseNumerics;
using NavfertyExcelAddIn.FindFormulaErrors;
using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.UnprotectWorkbook;

namespace NavfertyExcelAddIn
{
    public static class Registry
    {
        public static IContainer CreateContainer()
        {
            var builder = new ContainerBuilder();

            builder.RegisterType<DialogService>()
                .As<IDialogService>();

            builder.RegisterType<CellsUnmerger>()
                .As<ICellsUnmerger>()
                .EnableInterfaceInterceptors()
                .InterceptedBy(typeof(ExceptionLogger));

            builder.RegisterType<WbUnprotector>()
                .As<IWbUnprotector>()
                .EnableInterfaceInterceptors()
                .InterceptedBy(typeof(ExceptionLogger));

            builder.RegisterType<ErrorFinder>()
                .As<IErrorFinder>()
                .EnableInterfaceInterceptors()
                .InterceptedBy(typeof(ExceptionLogger));

            builder.RegisterType<NumericParser>()
                .As<INumericParser>()
                .EnableInterfaceInterceptors()
                .InterceptedBy(typeof(ExceptionLogger));

            builder.RegisterType<ExceptionLogger>();


            return builder.Build();
        }
    }
}
