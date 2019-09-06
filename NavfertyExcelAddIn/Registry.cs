using Autofac;
using Autofac.Extras.DynamicProxy;
using NavfertyExcelAddIn.ParseNumerics;
using NavfertyExcelAddIn.Commons;

namespace NavfertyExcelAddIn
{
    public static class Registry
    {
        public static IContainer CreateContainer()
        {
            var builder = new ContainerBuilder();

            builder.RegisterType<ErrorFinder>()
                .As<IErrorFinder>()
                .EnableInterfaceInterceptors()
                .InterceptedBy(typeof(ExceptionLogger));

            builder.RegisterType<NumericParser>()
                .As<INumericParser>()
                .EnableInterfaceInterceptors()
                .InterceptedBy(typeof(ExceptionLogger));

            builder.Register(c => new ExceptionLogger());


            return builder.Build();
        }
    }
}
