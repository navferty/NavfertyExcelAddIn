using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace NavfertyExcelAddIn.UnitTests
{
    [TestClass]
    public class TestUnitTest
    {
        [TestMethod]
        public void TestMethod()
        {
            Assert.AreEqual(2, 1 + 1);

            Assert.ThrowsException<InvalidOperationException>(() => ThrowEx());
        }

        private void ThrowEx() => throw new InvalidOperationException("Error");
    }
}
