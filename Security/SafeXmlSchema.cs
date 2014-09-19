//---------------------------------------------------------------------
// <copyright file="SafeXmlSchema.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Data
{
    using System.IO;
    using System.Xml;
    using System.Xml.Schema;

    /// <summary>
    /// XmlSchema with protection against DTD parsing in read overloads.
    /// </summary>
    internal class SafeXmlSchema : XmlSchema
    {
        #region Members
        /// <summary>
        /// Safe xml reader settings.
        /// </summary>
        private static XmlReaderSettings defaultSettings = new XmlReaderSettings()
        {
            ProhibitDtd = true,
            XmlResolver = null
        };
        #endregion

        #region Methods
        /// <summary>
        /// Reads an XML Schema from the supplied stream.
        /// </summary>
        /// <param name="stream">The supplied data stream.</param>
        /// <param name="validationEventHandler">The validation event handler that receives information about the XML Schema syntax errors.</param>
        /// <returns>The XmlSchema object representing the XML Schema.</returns>
        public static new XmlSchema Read(Stream stream, ValidationEventHandler validationEventHandler)
        {
            using (XmlReader xr = XmlReader.Create(stream, SafeXmlSchema.defaultSettings))
            {
                return XmlSchema.Read(xr, validationEventHandler);
            }
        }

        /// <summary>
        /// Reads an XML Schema from the supplied TextReader.
        /// </summary>
        /// <param name="reader">The TextReader containing the XML Schema to read.</param>
        /// <param name="validationEventHandler">The validation event handler that receives information about the XML Schema syntax errors.</param>
        /// <returns>The XmlSchema object representing the XML Schema.</returns>
        public static new XmlSchema Read(TextReader reader, ValidationEventHandler validationEventHandler)
        {
            using (XmlReader xr = XmlReader.Create(reader, SafeXmlSchema.defaultSettings))
            {
                return XmlSchema.Read(xr, validationEventHandler);
            }
        }

        /// <summary>
        /// Reads an XML Schema from the supplied XmlReader.
        /// </summary>
        /// <param name="reader">The XmlReader containing the XML Schema to read.</param>
        /// <param name="validationEventHandler">The validation event handler that receives information about the XML Schema syntax errors.</param>
        /// <returns>The XmlSchema object representing the XML Schema.</returns>
        public static new XmlSchema Read(XmlReader reader, ValidationEventHandler validationEventHandler)
        {
            // we need to check to see if the reader is configured properly
            if (reader.Settings != null)
            {
                if (reader.Settings.ProhibitDtd != true)
                {
                    throw new XmlDtdException();
                }
            }

            return XmlSchema.Read(reader, validationEventHandler);
        }
        #endregion
    }
}