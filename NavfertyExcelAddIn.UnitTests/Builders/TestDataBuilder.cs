using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NavfertyExcelAddIn.UnitTests.Builders
{
    public abstract class TestDataBuilder<TItem>
        where TItem : class
    {
        public abstract TItem Build();
    }
}
