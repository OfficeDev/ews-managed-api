/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
{
    using System.IO;
    using System.Xml;
    using System.Xml.XPath;

    /// <summary>
    /// Factory methods to safely instantiate XXE vulnerable object.
    /// </summary>
    internal class SafeXmlFactory
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

        #region XmlTextReader
        /// <summary>
        /// Initializes a new instance of the XmlTextReader class with the specified stream.
        /// </summary>
        /// <param name="stream">The stream containing the XML data to read.</param>
        /// <returns>A new instance of the XmlTextReader class.</returns>
        public static XmlTextReader CreateSafeXmlTextReader(Stream stream)
        {
            XmlTextReader xtr = new XmlTextReader(stream);
            xtr.ProhibitDtd = true;
            xtr.XmlResolver = null;
            return xtr;
        }

        /// <summary>
        /// Initializes a new instance of the XmlTextReader class with the specified file.
        /// </summary>
        /// <param name="url">The URL for the file containing the XML data. The BaseURI is set to this value.</param>
        /// <returns>A new instance of the XmlTextReader class.</returns>
        public static XmlTextReader CreateSafeXmlTextReader(string url)
        {
            XmlTextReader xtr = new XmlTextReader(url);
            xtr.ProhibitDtd = true;
            xtr.XmlResolver = null;
            return xtr;
        }

        /// <summary>
        /// Initializes a new instance of the XmlTextReader class with the specified TextReader.
        /// </summary>
        /// <param name="input">The TextReader containing the XML data to read.</param>
        /// <returns>A new instance of the XmlTextReader class.</returns>
        public static XmlTextReader CreateSafeXmlTextReader(TextReader input)
        {
            XmlTextReader xtr = new XmlTextReader(input);
            xtr.ProhibitDtd = true;
            xtr.XmlResolver = null;
            return xtr;
        }

        /// <summary>
        /// Initializes a new instance of the XmlTextReader class with the specified stream and XmlNameTable.
        /// </summary>
        /// <param name="input">The stream containing the XML data to read.</param>
        /// <param name="nt">The XmlNameTable to use.</param>
        /// <returns>A new instance of the XmlTextReader class.</returns>
        public static XmlTextReader CreateSafeXmlTextReader(Stream input, XmlNameTable nt)
        {
            XmlTextReader xtr = new XmlTextReader(input, nt);
            xtr.ProhibitDtd = true;
            xtr.XmlResolver = null;
            return xtr;
        }

        /// <summary>
        /// Initializes a new instance of the XmlTextReader class with the specified URL and stream.
        /// </summary>
        /// <param name="url">The URL to use for resolving external resources. The BaseURI is set to this value.</param>
        /// <param name="input">The stream containing the XML data to read.</param>
        /// <returns>A new instance of the XmlTextReader class.</returns>
        public static XmlTextReader CreateSafeXmlTextReader(string url, Stream input)
        {
            XmlTextReader xtr = new XmlTextReader(url, input);
            xtr.ProhibitDtd = true;
            xtr.XmlResolver = null;
            return xtr;
        }

        /// <summary>
        /// Initializes a new instance of the XmlTextReader class with the specified TextReader.
        /// </summary>
        /// <param name="url">The URL to use for resolving external resources. The BaseURI is set to this value.</param>
        /// <param name="input">The TextReader containing the XML data to read.</param>
        /// <returns>A new instance of the XmlTextReader class.</returns>
        public static XmlTextReader CreateSafeXmlTextReader(string url, TextReader input)
        {
            XmlTextReader xtr = new XmlTextReader(url, input);
            xtr.ProhibitDtd = true;
            xtr.XmlResolver = null;
            return xtr;
        }

        /// <summary>
        /// Initializes a new instance of the XmlTextReader class with the specified file and XmlNameTable.
        /// </summary>
        /// <param name="url">The URL for the file containing the XML data to read.</param>
        /// <param name="nt">The XmlNameTable to use.</param>
        /// <returns>A new instance of the XmlTextReader class.</returns>
        public static XmlTextReader CreateSafeXmlTextReader(string url, XmlNameTable nt)
        {
            XmlTextReader xtr = new XmlTextReader(url, nt);
            xtr.ProhibitDtd = true;
            xtr.XmlResolver = null;
            return xtr;
        }

        /// <summary>
        /// Initializes a new instance of the XmlTextReader class with the specified TextReader.
        /// </summary>
        /// <param name="input">The TextReader containing the XML data to read.</param>
        /// <param name="nt">The XmlNameTable to use.</param>
        /// <returns>A new instance of the XmlTextReader class.</returns>
        public static XmlTextReader CreateSafeXmlTextReader(TextReader input, XmlNameTable nt)
        {
            XmlTextReader xtr = new XmlTextReader(input, nt);
            xtr.ProhibitDtd = true;
            xtr.XmlResolver = null;
            return xtr;
        }

        /// <summary>
        /// Initializes a new instance of the XmlTextReader class with the specified stream, XmlNodeType, and XmlParserContext.
        /// </summary>
        /// <param name="xmlFragment">The stream containing the XML fragment to parse.</param>
        /// <param name="fragType">The XmlNodeType of the XML fragment. This also determines what the fragment can contain.</param>
        /// <param name="context">The XmlParserContext in which the xmlFragment is to be parsed.</param>
        /// <returns>A new instance of the XmlTextReader class.</returns>
        public static XmlTextReader CreateSafeXmlTextReader(Stream xmlFragment, XmlNodeType fragType, XmlParserContext context)
        {
            XmlTextReader xtr = new XmlTextReader(xmlFragment, fragType, context);
            xtr.ProhibitDtd = true;
            xtr.XmlResolver = null;
            return xtr;
        }

        /// <summary>
        /// Initializes a new instance of the XmlTextReader class with the specified URL, stream and XmlNameTable.
        /// </summary>
        /// <param name="url">The URL to use for resolving external resources. The BaseURI is set to this value. If url is null, BaseURI is set to String.Empty.</param>
        /// <param name="input">The stream containing the XML data to read.</param>
        /// <param name="nt">The XmlNameTable to use.</param>
        /// <returns>A new instance of the XmlTextReader class.</returns>
        public static XmlTextReader CreateSafeXmlTextReader(string url, Stream input, XmlNameTable nt)
        {
            XmlTextReader xtr = new XmlTextReader(url, input, nt);
            xtr.ProhibitDtd = true;
            xtr.XmlResolver = null;
            return xtr;
        }

        /// <summary>
        /// Initializes a new instance of the XmlTextReader class with the specified URL, TextReader and XmlNameTable.
        /// </summary>
        /// <param name="url">The URL to use for resolving external resources. The BaseURI is set to this value. If url is null, BaseURI is set to String.Empty.</param>
        /// <param name="input">The TextReader containing the XML data to read.</param>
        /// <param name="nt">The XmlNameTable to use.</param>
        /// <returns>A new instance of the XmlTextReader class.</returns>
        public static XmlTextReader CreateSafeXmlTextReader(string url, TextReader input, XmlNameTable nt)
        {
            XmlTextReader xtr = new XmlTextReader(url, input, nt);
            xtr.ProhibitDtd = true;
            xtr.XmlResolver = null;
            return xtr;
        }

        /// <summary>
        /// Initializes a new instance of the XmlTextReader class with the specified string, XmlNodeType, and XmlParserContext.
        /// </summary>
        /// <param name="xmlFragment">The string containing the XML fragment to parse.</param>
        /// <param name="fragType">The XmlNodeType of the XML fragment. This also determines what the fragment string can contain.</param>
        /// <param name="context">The XmlParserContext in which the xmlFragment is to be parsed.</param>
        /// <returns>A new instance of the XmlTextReader class.</returns>
        public static XmlTextReader CreateSafeXmlTextReader(string xmlFragment, XmlNodeType fragType, XmlParserContext context)
        {
            XmlTextReader xtr = new XmlTextReader(xmlFragment, fragType, context);
            xtr.ProhibitDtd = true;
            xtr.XmlResolver = null;
            return xtr;
        }
        #endregion

        #region XPathDocument
        /// <summary>
        /// Initializes a new instance of the XPathDocument class from the XML data in the specified Stream object.
        /// </summary>
        /// <param name="stream">The Stream object that contains the XML data.</param>
        /// <returns>A new instance of the XPathDocument class.</returns>
        public static XPathDocument CreateXPathDocument(Stream stream)
        {
            using (XmlReader xr = XmlReader.Create(stream, SafeXmlFactory.defaultSettings))
            {
                return CreateXPathDocument(xr);
            }
        }

        /// <summary>
        /// Initializes a new instance of the XPathDocument class from the XML data in the specified file.
        /// </summary>
        /// <param name="uri">The path of the file that contains the XML data.</param>
        /// <returns>A new instance of the XPathDocument class.</returns>
        public static XPathDocument CreateXPathDocument(string uri)
        {
            using (XmlReader xr = XmlReader.Create(uri, SafeXmlFactory.defaultSettings))
            {
                return CreateXPathDocument(xr);
            }
        }

        /// <summary>
        /// Initializes a new instance of the XPathDocument class from the XML data that is contained in the specified TextReader object.
        /// </summary>
        /// <param name="textReader">The TextReader object that contains the XML data.</param>
        /// <returns>A new instance of the XPathDocument class.</returns>
        public static XPathDocument CreateXPathDocument(TextReader textReader)
        {
            using (XmlReader xr = XmlReader.Create(textReader, SafeXmlFactory.defaultSettings))
            {
                return CreateXPathDocument(xr);
            }
        }

        /// <summary>
        /// Initializes a new instance of the XPathDocument class from the XML data that is contained in the specified XmlReader object.
        /// </summary>
        /// <param name="reader">The XmlReader object that contains the XML data.</param>
        /// <returns>A new instance of the XPathDocument class.</returns>
        public static XPathDocument CreateXPathDocument(XmlReader reader)
        {
            // we need to check to see if the reader is configured properly
            if (reader.Settings != null)
            {
                if (reader.Settings.ProhibitDtd != true)
                {
                    throw new XmlDtdException();
                }
            }

            return new XPathDocument(reader);
        }

        /// <summary>
        /// Initializes a new instance of the XPathDocument class from the XML data in the file specified with the white space handling specified.
        /// </summary>
        /// <param name="uri">The path of the file that contains the XML data.</param>
        /// <param name="space">An XmlSpace object.</param>
        /// <returns>A new instance of the XPathDocument class.</returns>
        public static XPathDocument CreateXPathDocument(string uri, XmlSpace space)
        {
            using (XmlReader xr = XmlReader.Create(uri, SafeXmlFactory.defaultSettings))
            {
                return CreateXPathDocument(xr, space);
            }
        }

        /// <summary>
        /// Initializes a new instance of the XPathDocument class from the XML data that is contained in the specified XmlReader object with the specified white space handling.
        /// </summary>
        /// <param name="reader">The XmlReader object that contains the XML data.</param>
        /// <param name="space">An XmlSpace object.</param>
        /// <returns>A new instance of the XPathDocument class.</returns>
        public static XPathDocument CreateXPathDocument(XmlReader reader, XmlSpace space)
        {
            // we need to check to see if the reader is configured properly
            if (reader.Settings != null)
            {
                if (reader.Settings.ProhibitDtd != true)
                {
                    throw new XmlDtdException();
                }
            }

            return new XPathDocument(reader, space);
        }
        #endregion
    }
}