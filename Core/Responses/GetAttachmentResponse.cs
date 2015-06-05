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
    using System.Xml;

    /// <summary>
    /// Represents the response to an individual attachment retrieval request.
    /// </summary>
    public sealed class GetAttachmentResponse : ServiceResponse
    {
        private Attachment attachment;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetAttachmentResponse"/> class.
        /// </summary>
        /// <param name="attachment">The attachment.</param>
        internal GetAttachmentResponse(Attachment attachment)
            : base()
        {
            this.attachment = attachment;
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.Attachments);
            if (!reader.IsEmptyElement)
            {
                reader.Read(XmlNodeType.Element);

                if (this.attachment == null)
                {
                    if (string.Equals(reader.LocalName, XmlElementNames.FileAttachment, StringComparison.OrdinalIgnoreCase))
                    {
                        this.attachment = new FileAttachment(reader.Service);
                    }
                    else if (string.Equals(reader.LocalName, XmlElementNames.ItemAttachment, StringComparison.OrdinalIgnoreCase))
                    {
                        this.attachment = new ItemAttachment(reader.Service);
                    }
                }

                if (this.attachment != null)
                {
                    this.attachment.LoadFromXml(reader, reader.LocalName);
                }

                reader.ReadEndElement(XmlNamespace.Messages, XmlElementNames.Attachments);
            }
        }

        /// <summary>
        /// Gets the attachment that was retrieved.
        /// </summary>
        public Attachment Attachment
        {
            get { return this.attachment; }
        }
    }
}