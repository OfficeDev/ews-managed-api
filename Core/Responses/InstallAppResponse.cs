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
    using System.Collections.ObjectModel;
    using System.IO;
    using System.Xml;

    /// <summary>
    /// Represents the response to a InstallApp operation.
    /// Today this class doesn't add extra functionality. Keep this class here so future
    /// we can return extension info up-on installation complete. 
    /// </summary>
    internal sealed class InstallAppResponse : ServiceResponse
    {
        private bool? wasFirstInstall;

        /// <summary>
        /// Initializes a new instance of the <see cref="InstallAppResponse"/> class.
        /// </summary>
        public InstallAppResponse()
            : base()
        {
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            if (this.ErrorCode == ServiceError.NoError && reader.IsStartElement(XmlNamespace.NotSpecified, XmlElementNames.WasFirstInstall))
            {
                this.wasFirstInstall = reader.ReadElementValue<bool>(XmlNamespace.NotSpecified, XmlElementNames.WasFirstInstall);
            }
        }

        /// <summary>
        /// Was this first install
        /// </summary>
        public bool? WasFirstInstall
        {
            get { return this.wasFirstInstall; }
        }
    }
}