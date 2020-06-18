using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Schema;
using System.Windows.Forms;
using System.IO;

namespace LeadsExtractor
{
    class clsSValidator
    {

        private string m_sXMLFileName;
        private string m_sSchemaFileName;      
        
        static int numErrors = 0;
        static string msgError = "";

        public clsSValidator(string sXMLFileName, string sSchemaFileName)
        {

            m_sXMLFileName = sXMLFileName;
            m_sSchemaFileName = sSchemaFileName;

 
        }

        public bool ValidateXMLFile()
        {
            if (Validate(m_sXMLFileName, m_sSchemaFileName) != true)
            {
                MessageBox.Show("XML file doesn't matches the schema! Please verify the input format!");
                return false;
            }

            return true;
        }

        private static void ErrorHandler(object sender, ValidationEventArgs args)
        {
            msgError = msgError + "\r\n" + args.Message;
            numErrors++;
        }

        private static bool Validate(string xmlFilename, string xsdFilename)
        {
            return Validate(GetFileStream(xmlFilename), GetFileStream(xsdFilename));
        }

        private static void ClearErrorMessage()
        {
            msgError = "";
            numErrors = 0;
        }

        // returns a stream of the contents of the given filename
        private static Stream GetFileStream(string filename)
        {
            try
            {
                return new FileStream(filename, FileMode.Open, FileAccess.Read);
            }
            catch
            {
                return null;
            }
        }

        private static bool Validate(Stream xml, Stream xsd)
        {
            ClearErrorMessage();
            try
            {
                XmlTextReader tr = new XmlTextReader(xsd);
                XmlSchemaSet schema = new XmlSchemaSet();
                schema.Add(null, tr);

                XmlReaderSettings settings = new XmlReaderSettings();
                settings.ValidationType = ValidationType.Schema;
                settings.Schemas.Add(schema);
                settings.ValidationFlags |= XmlSchemaValidationFlags.ReportValidationWarnings;
                settings.ValidationEventHandler += new ValidationEventHandler(ErrorHandler);
                XmlReader reader = XmlReader.Create(xml, settings);

                // Validate XML data
                while (reader.Read())
                    ;
                reader.Close();                

                // exception if validation failed
                if (numErrors > 0)
                    throw new Exception(msgError);

                tr.Close();
                xml.Close();
                xml.Dispose();
                xsd.Close();
                xsd.Dispose();

                return true;
            }
            catch
            {
                msgError = "Validation failed\r\n" + msgError;
                return false;
            }
        }   


    
    }


}
