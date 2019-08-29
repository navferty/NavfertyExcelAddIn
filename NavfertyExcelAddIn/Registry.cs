using Autofac;
using Autofac.Extras.DynamicProxy;
using NavfertyExcelAddIn.ParseNumerics;

namespace NavfertyExcelAddIn
{
    public static class Registry
    {
        public static IContainer CreateContainer()
        {
            var builder = new ContainerBuilder();



            builder.RegisterType<NumericParser>()
                .As<INumericParser>()
                .EnableInterfaceInterceptors();

            builder.Register(c => new ExceptionLogger());


            return builder.Build();
        }
    }
}
