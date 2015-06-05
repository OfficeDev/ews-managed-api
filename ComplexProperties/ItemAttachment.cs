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
    /// Represents an item attachment.
    /// </summary>
    public class ItemAttachment : Attachment
    {
        /// <summary>
        /// The item associated with the attachment.
        /// </summary>
        private Item item;

        /// <summary>
        /// Initializes a new instance of the <see cref="ItemAttachment"/> class.
        /// </summary>
        /// <param name="owner">The owner of the attachment.</param>
        internal ItemAttachment(Item owner)
            : base(owner)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ItemAttachment"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal ItemAttachment(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Gets the item associated with the attachment.
        /// </summary>
        public Item Item
        {
            get
            {
                return this.item;
            }

            internal set
            {
                this.ThrowIfThisIsNotNew();

                if (this.item != null)
                {
                    this.item.OnChange -= this.ItemChanged;
                }
                this.item = value;
                if (this.item != null)
                {
                    this.item.OnChange += this.ItemChanged;
                }
            }
        }

        /// <summary>
        /// Implements the OnChange event handler for the item associated with the attachment.
        /// </summary>
        /// <param name="serviceObject">The service object that triggered the OnChange event.</param>
        private void ItemChanged(ServiceObject serviceObject)
        {
            if (this.Owner != null)
            {
                this.Owner.PropertyBag.Changed();
            }
        }

        /// <summary>
        /// Obtains EWS XML element name for this object.
        /// </summary>
        /// <returns>The XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.ItemAttachment;
        }

        /// <summary>
        /// Tries to read the element at the current position of the reader.
        /// </summary>
        /// <param name="reader">The reader to read the element from.</param>
        /// <returns>True if the element was read, false otherwise.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            bool result = base.TryReadElementFromXml(reader);

            if (!result)
            {
                this.item = EwsUtilities.CreateItemFromXmlElementName(this, reader.LocalName);

                if (this.item != null)
                {
                    this.item.LoadFromXml(reader, true /* clearPropertyBag */);
                }
            }

            return result;
        }

        /// <summary>
        /// For ItemAttachment, AttachmentId and Item should be patched. 
        /// </summary>
        /// <param name="reader"></param>
        /// <returns></returns>
        internal override bool TryReadElementFromXmlToPatch(EwsServiceXmlReader reader)
        {
            // update the attachment id.
            base.TryReadElementFromXml(reader);

            reader.Read();
            Type itemClass = EwsUtilities.GetItemTypeFromXmlElementName(reader.LocalName);

            if (itemClass != null)
            {
                if (this.item == null || this.item.GetType() != itemClass)
                {
                    throw new ServiceLocalException(Strings.AttachmentItemTypeMismatch);
                }

                this.item.LoadFromXml(reader, false /* clearPropertyBag */);
                return true;
            }

            return false;
        }

        /// <summary>
        /// Writes the properties of this object as XML elements.
        /// </summary>
        /// <param name="writer">The writer to write the elements to.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            base.WriteElementsToXml(writer);

            this.Item.WriteToXml(writer);
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        /// <param name="attachmentIndex">Index of this attachment.</param>
        internal override void Validate(int attachmentIndex)
        {
            if (string.IsNullOrEmpty(this.Name))
            {
                throw new ServiceValidationException(string.Format(Strings.ItemAttachmentMustBeNamed, attachmentIndex));
            }

            // Recurse through any items attached to item attachment.
            this.Item.Attachments.Validate();
        }

        /// <summary>
        /// Loads this attachment.
        /// </summary>
        /// <param name="additionalProperties">The optional additional properties to load.</param>
        public void Load(params PropertyDefinitionBase[] additionalProperties)
        {
            this.InternalLoad(
                null /* bodyType */,
                additionalProperties);
        }

        /// <summary>
        /// Loads this attachment.
        /// </summary>
        /// <param name="additionalProperties">The optional additional properties to load.</param>
        public void Load(IEnumerable<PropertyDefinitionBase> additionalProperties)
        {
            this.InternalLoad(
                null /* bodyType */,
                additionalProperties);
        }

        /// <summary>
        /// Loads this attachment.
        /// </summary>
        /// <param name="bodyType">The body type to load.</param>
        /// <param name="additionalProperties">The optional additional properties to load.</param>
        public void Load(BodyType bodyType, params PropertyDefinitionBase[] additionalProperties)
        {
            this.InternalLoad(
                bodyType,
                additionalProperties);
        }

        /// <summary>
        /// Loads this attachment.
        /// </summary>
        /// <param name="bodyType">The body type to load.</param>
        /// <param name="additionalProperties">The optional additional properties to load.</param>
        public void Load(BodyType bodyType, IEnumerable<PropertyDefinitionBase> additionalProperties)
        {
            this.InternalLoad(
                bodyType,
                additionalProperties);
        }
    }
}