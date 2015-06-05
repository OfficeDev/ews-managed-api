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
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the response to an individual Id conversion operation.
    /// </summary>
    public sealed class ConvertIdResponse : ServiceResponse
    {
        private AlternateIdBase convertedId;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConvertIdResponse"/> class.
        /// </summary>
        internal ConvertIdResponse()
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

            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.AlternateId);

            string alternateIdClass = reader.ReadAttributeValue(XmlNamespace.XmlSchemaInstance, XmlAttributeNames.Type);

            int aliasSeparatorIndex = alternateIdClass.IndexOf(':');

            if (aliasSeparatorIndex > -1)
            {
                alternateIdClass = alternateIdClass.Substring(aliasSeparatorIndex + 1);
            }

            // Alternate Id classes are responsible fro reading the AlternateId end element when necessary
            switch (alternateIdClass)
            {
                case AlternateId.SchemaTypeName:
                    this.convertedId = new AlternateId();
                    break;
                case AlternatePublicFolderId.SchemaTypeName:
                    this.convertedId = new AlternatePublicFolderId();
                    break;
                case AlternatePublicFolderItemId.SchemaTypeName:
                    this.convertedId = new AlternatePublicFolderItemId();
                    break;
                default:
                    EwsUtilities.Assert(
                        false,
                        "ConvertIdResponse.ReadElementsFromXml",
                        string.Format("Unknown alternate Id class: {0}", alternateIdClass));
                    break;
            }

            this.convertedId.LoadAttributesFromXml(reader);

            reader.ReadEndElementIfNecessary(XmlNamespace.Messages, XmlElementNames.AlternateId);
        }

        /// <summary>
        /// Gets the converted Id.
        /// </summary>
        public AlternateIdBase ConvertedId
        {
            get { return this.convertedId; }
        }
    }
}