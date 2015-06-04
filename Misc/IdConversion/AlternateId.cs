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
    /// <summary>
    /// Represents an Id expressed in a specific format.
    /// </summary>
    public class AlternateId : AlternateIdBase
    {
        /// <summary>
        /// Name of schema type used for AlternateId.
        /// </summary>
        internal const string SchemaTypeName = "AlternateIdType";

        /// <summary>
        /// Initializes a new instance of the <see cref="AlternateId"/> class.
        /// </summary>
        public AlternateId()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AlternateId"/> class.
        /// </summary>
        /// <param name="format">The format the Id is expressed in.</param>
        /// <param name="id">The Id.</param>
        /// <param name="mailbox">The SMTP address of the mailbox that the Id belongs to.</param>
        public AlternateId(
            IdFormat format,
            string id,
            string mailbox)
            : base(format)
        {
            this.UniqueId = id;
            this.Mailbox = mailbox;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AlternateId"/> class.
        /// </summary>
        /// <param name="format">The format the Id is expressed in.</param>
        /// <param name="id">The Id.</param>
        /// <param name="mailbox">The SMTP address of the mailbox that the Id belongs to.</param>
        /// <param name="isArchive">Primary (false) or archive (true) mailbox.</param>
        public AlternateId(
            IdFormat format,
            string id,
            string mailbox,
            bool isArchive)
            : base(format)
        {
            this.UniqueId = id;
            this.Mailbox = mailbox;
            this.IsArchive = isArchive;
        }

        /// <summary>
        /// Gets or sets the Id.
        /// </summary>
        public string UniqueId
        {
            get; set;
        }

        /// <summary>
        /// Gets or sets the mailbox to which the Id belongs.
        /// </summary>
        public string Mailbox
        {
            get; set;
        }

        /// <summary>
        /// Gets or sets the type (primary or archive) mailbox to which the Id belongs.
        /// </summary>
        public bool IsArchive
        {
            get; set;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.AlternateId;
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            base.WriteAttributesToXml(writer);

            writer.WriteAttributeValue(XmlAttributeNames.Id, this.UniqueId);
            writer.WriteAttributeValue(XmlAttributeNames.Mailbox, this.Mailbox);

            // this is optional attribute will default false so we will write
            // it only if it is true
            if (this.IsArchive)
            {
                writer.WriteAttributeValue(XmlAttributeNames.IsArchive, true);
            }
        }

        /// <summary>
        /// Loads the attributes from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void LoadAttributesFromXml(EwsServiceXmlReader reader)
        {
            base.LoadAttributesFromXml(reader);

            this.UniqueId = reader.ReadAttributeValue(XmlAttributeNames.Id);
            this.Mailbox = reader.ReadAttributeValue(XmlAttributeNames.Mailbox);

            // optional attribute: defaulting to false if not present
            string isArchive = reader.ReadAttributeValue(XmlAttributeNames.IsArchive);
            if (!string.IsNullOrEmpty(isArchive))
            {
                this.IsArchive = reader.ReadAttributeValue<bool>(XmlAttributeNames.IsArchive);
            }
            else
            {
                this.IsArchive = false;
            }
        }

        /// <summary>
        /// Validate this instance.
        /// </summary>
        internal override void InternalValidate()
        {
            EwsUtilities.ValidateParam(this.Mailbox, "mailbox");
        }
    }
}