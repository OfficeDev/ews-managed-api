#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

//-----------------------------------------------------------------------
// <summary>Defines the EwsServiceXmlReader class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// Represents an xml reader used by the ExchangeService to parse multi-response streams, 
    /// such as GetStreamingEvents. 
    /// </summary>
    /// <remarks>
    /// Necessary because the basic EwsServiceXmlReader does not 
    /// use normalization, and in order to turn normalization off, it is 
    /// necessary to use an XmlTextReader, which does not allow the ConformanceLevel.Auto that
    /// a multi-response stream requires.
    /// If ever there comes a time we need to deal with multi-response streams with user-generated
    /// content, we will need to tackle that parsing problem separately.
    /// </remarks>
    internal class EwsServiceMultiResponseXmlReader : EwsServiceXmlReader
    {
        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="EwsServiceMultiResponseXmlReader"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="service">The service.</param>
        private EwsServiceMultiResponseXmlReader(Stream stream, ExchangeService service)
            : base(stream, service)
        {
        }

        /// <summary>
        /// Creates a new instance of the <see cref="EwsServiceMultiResponseXmlReader"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="service">The service.</param>
        /// <returns>an instance of EwsServiceMultiResponseXmlReader wrapped around the input stream.</returns>
        internal static EwsServiceMultiResponseXmlReader Create(Stream stream, ExchangeService service)
        {
            EwsServiceMultiResponseXmlReader reader = new EwsServiceMultiResponseXmlReader(stream, service);

            return reader;
        }
        #endregion

        /// <summary>
        /// Creates the XML reader.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns>An XML reader to use.</returns>
        private static XmlReader CreateXmlReader(Stream stream)
        {
            // The ProhibitDtd property is used to indicate whether XmlReader should process DTDs or not. By default, 
            // it will do so. EWS doesn't use DTD references so we want to turn this off. Also, the XmlResolver property is
            // set to an instance of XmlUrlResolver by default. We don't want XmlTextReader to try to resolve this DTD reference 
            // so we disable the XmlResolver as well.
            XmlReaderSettings settings = new XmlReaderSettings()
            {
                ConformanceLevel = ConformanceLevel.Auto,
                ProhibitDtd = true,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                IgnoreWhitespace = true,
                XmlResolver = null
            };

            return XmlReader.Create(stream, settings);
        }

        /// <summary>
        /// Initializes the XML reader.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns>An XML reader to use.</returns>
        protected override XmlReader InitializeXmlReader(Stream stream)
        {
            return CreateXmlReader(stream);
        }
    }
}
