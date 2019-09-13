using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Schema;
using NLog;

namespace NavfertyExcelAddIn.XmlTools
{
    public class XsdSchemaValidator : IXsdSchemaValidator
    {
        // private readonly Logger logger = LogManager.GetCurrentClassLogger();

        public IReadOnlyCollection<string> Validate(string xmlFilename, IReadOnlyCollection<string> schemas)
        {
            var settings = new XmlReaderSettings { ValidationType = ValidationType.Schema };
            foreach (var schema in schemas)
            {
                var loadedSchema = LoadSchema(schema);
                settings.Schemas.Add(loadedSchema);
            }

            var validationErrors = new List<string>();

            settings.ValidationEventHandler += (object sender, ValidationEventArgs e)
                => validationErrors.Add(TransformError((XmlReader)sender, e));

            using (var xmlStream = File.OpenRead(xmlFilename))
            using (var xmlFile = XmlReader.Create(xmlStream, settings))
            {
                while (xmlFile.Read())
                {
                    // do nothing
                }
            }

            return validationErrors.ToArray();
        }

        private XmlSchema LoadSchema(string xsdFilename)
        {
            var errors = new List<string>();

            XmlSchema schema;
            using (var fs = File.OpenRead(xsdFilename))
            {
                schema = XmlSchema.Read(fs, (object sender, ValidationEventArgs e)
                    => errors.Add(TransformError((XmlReader)sender, e).ToString()));
            }

            if (errors.Any())
            {
                var message = "Error occured while adding schema:\r\n" + string.Join(", ", errors);
                throw new InvalidOperationException(message);
            }

            return schema;
        }

        private string TransformError(XmlReader reader, ValidationEventArgs e) =>
            $"{e.Severity.ToString()}. {reader.Name}: '{reader.Value}': {e.Message}";
    }
}
