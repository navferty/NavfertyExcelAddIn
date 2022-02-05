using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

using Microsoft.Xml.XMLGen;

using Navferty.Common;

using NLog;

namespace NavfertyExcelAddIn.XmlTools
{
    public class XmlSampleCreator : IXmlSampleCreator
    {
        private readonly IDialogService dialogService;
        private readonly Logger logger = LogManager.GetCurrentClassLogger();

        public XmlSampleCreator(IDialogService dialogService)
        {
            this.dialogService = dialogService;
        }

        public void CreateSampleXml()
        {
            var xsdFileName = dialogService.AskForFiles(false, FileType.Xsd).FirstOrDefault();
            if (string.IsNullOrEmpty(xsdFileName))
            {
                logger.Error("Empty file path");
                return;
            }

            var targetFileName = dialogService.AskFileNameSaveAs("sample.xml", FileType.Xml);
            if (string.IsNullOrEmpty(targetFileName))
            {
                logger.Error("Empty target file path");
                return;
            }

            var sampleGenerator = new XmlSampleGenerator(xsdFileName, XmlQualifiedName.Empty);

            using (var fileStream = new FileStream(targetFileName, FileMode.CreateNew))
            using (var xmlTextWriter = new XmlTextWriter(fileStream, Encoding.UTF8))
            {
                xmlTextWriter.Formatting = Formatting.Indented;
                sampleGenerator.WriteXml(xmlTextWriter);
            }
        }
    }
}
