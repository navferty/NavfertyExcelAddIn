using System;
using System.Collections;
using System.Linq;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Moq;

using NavfertyExcelAddIn.DataValidation;

using Range = Microsoft.Office.Interop.Excel.Range;

namespace NavfertyExcelAddIn.UnitTests.DataValidation
{
    [TestClass]
    public class CellsValueValidatorTests : TestsBase
    {
        private Mock<Range> selection;
        private Mock<IValidator> validator;

        private CellsValueValidator cellsValidator;

        [TestInitialize]
        public void BeforeEachTest()
        {
            selection = GetRangeStub();

            validator = new Mock<IValidator>(MockBehavior.Strict);

            var validatorFactory = new Mock<IValidatorFactory>(MockBehavior.Strict);
            validatorFactory.Setup(x => x.CreateValidator(ValidationType.Xml)).Returns(validator.Object);

            cellsValidator = new CellsValueValidator(validatorFactory.Object);
        }

        [TestMethod]
        public void Validate_All_Success()
        {
            validator.Setup(x => x.CheckValue(It.IsAny<object>())).Returns(ValidationResult.Success);
            selection
                .As<IEnumerable>()
                .Setup(x => x.GetEnumerator())
                .Returns(Enumerable.Range(0, 10).Select(z => GetCellStub(z).Object).GetEnumerator());

            var result = cellsValidator.Validate(selection.Object, ValidationType.Xml);

            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void Validate_All_Fail()
        {
            validator.Setup(x => x.CheckValue(It.IsAny<object>())).Returns(ValidationResult.Fail(string.Empty));
            selection
                .As<IEnumerable>()
                .Setup(x => x.GetEnumerator())
                .Returns(Enumerable.Range(0, 10).Select(z => GetCellStub(z).Object).GetEnumerator());

            var result = cellsValidator.Validate(selection.Object, ValidationType.Xml);

            Assert.AreEqual(10, result.Count);
        }

        [TestMethod]
        public void Validate_Success_And_Fail()
        {
            var cells = new[]
            {
                GetCellStub(true),
                GetCellStub(true),
                GetCellStub(false),
                GetCellStub("")
            };
            selection.As<IEnumerable>().Setup(x => x.GetEnumerator()).Returns(cells.Select(c => c.Object).GetEnumerator());

            validator
                .Setup(x => x.CheckValue(It.IsAny<object>()))
                .Returns((bool x) => x ? ValidationResult.Success : ValidationResult.Fail("Fail =("));

            var result = cellsValidator.Validate(selection.Object, ValidationType.Xml);

            Assert.AreEqual(1, result.Count);
            validator.Verify(x => x.CheckValue(It.IsAny<object>()), Times.Exactly(3));
        }

        [TestMethod]
        public void Validate_NoCells()
        {
            selection.As<IEnumerable>().Setup(x => x.GetEnumerator()).Returns(Array.Empty<Range>().GetEnumerator());

            var result = cellsValidator.Validate(selection.Object, ValidationType.Xml);

            Assert.AreEqual(0, result.Count);
        }
    }
}
