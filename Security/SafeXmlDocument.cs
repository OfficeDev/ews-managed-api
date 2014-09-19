//---------------------------------------------------------------------
// <copyright file="SafeXmlDocument.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.IO;
    using System.Xml;

    /// <summary>
    /// XmlDocument that does not allow DTD parsing.
    /// </summary>
    internal class SafeXmlDocument : XmlDocument
    {
        #region Members
        /// <summary>
        /// Xml settings object.
        /// </summary>
        private XmlReaderSettings settings = new XmlReaderSettings()
        {
            ProhibitDtd = true,
            XmlResolver = null
        };
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the SafeXmlDocument class.
        /// </summary>
        public SafeXmlDocument()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the SafeXmlDocument class with the specified XmlImplementation.
        /// </summary>
        /// <remarks>Not supported do to no use within exchange dev code.</remarks>
        /// <param name="imp">The XmlImplementation to use.</param>
        public SafeXmlDocument(XmlImplementation imp)
        {
            throw new NotSupportedException("Not supported");
        }

        /// <summary>
        /// Initializes a new instance of the SafeXmlDocument class with the specified XmlNameTable.
        /// </summary>
        /// <param name="nt">The XmlNameTable to use.</param>
        public SafeXmlDocument(XmlNameTable nt)
            : base(nt)
        {
        }
        #endregion

        #region Methods
        /// <summary>
        /// Loads the XML document from the specified stream.
        /// </summary>
        /// <param name="inStream">The stream containing the XML document to load.</param>
        public override void Load(Stream inStream)
        {
            using (XmlReader reader = XmlReader.Create(inStream, this.settings))
            {
                this.Load(reader);
            }
        }

        /// <summary>
        /// Loads the XML document from the specified URL.
        /// </summary>
        /// <param name="filename">URL for the file containing the XML document to load. The URL can be either a local file or an HTTP URL (a Web address).</param>
        public override void Load(string filename)
        {
            using (XmlReader reader = XmlReader.Create(filename, this.settings))
            {
                this.Load(reader);
            }
        }

        /// <summary>
        /// Loads the XML document from the specified TextReader.
        /// </summary>
        /// <param name="txtReader">The TextReader used to feed the XML data into the document.</param>
        public override void Load(TextReader txtReader)
        {
            using (XmlReader reader = XmlReader.Create(txtReader, this.settings))
            {
                this.Load(reader);
            }
        }

        /// <summary>
        /// Loads the XML document from the specified XmlReader.
        /// </summary>
        /// <param name="reader">The XmlReader used to feed the XML data into the document.</param>
        public override void Load(XmlReader reader)
        {
            // we need to check to see if the reader is configured properly
            if (reader.Settings != null)
            {
                if (reader.Settings.ProhibitDtd != true)
                {
                    throw new XmlDtdException();
                }
            }

            try
            {
                base.Load(reader);
            }
            catch (XmlException x)
            {
                if (x.Message.StartsWith("For security reasons DTD is prohibited in this XML document.", StringComparison.OrdinalIgnoreCase))
                {
                    throw new XmlDtdException();
                }
            }
        }

        /// <summary>
        /// Loads the XML document from the specified string.
        /// </summary>
        /// <param name="xml">String containing the XML document to load.</param>
        public override void LoadXml(string xml)
        {
            using (XmlReader reader = XmlReader.Create(new StringReader(xml), this.settings))
            {
                base.Load(reader);
            }
        }
        #endregion
    }
}