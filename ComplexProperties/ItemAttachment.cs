// ---------------------------------------------------------------------------
// <copyright file="ItemAttachment.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ItemAttachment class.</summary>
//-----------------------------------------------------------------------

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
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service"></param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            base.LoadFromJson(jsonProperty, service);

            if (jsonProperty.ContainsKey(XmlElementNames.Item))
            {
                JsonObject jsonItem = jsonProperty.ReadAsJsonObject(XmlElementNames.Item);

                // skip this - "Item" : null
                if (jsonItem != null)
                {
                    this.item = EwsUtilities.CreateItemFromXmlElementName(this, jsonItem.ReadTypeString());

                    if (this.item != null)
                    {
                        this.item.LoadFromJson(jsonItem, service, true /* clearPropertyBag */);
                    }
                }
            }
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
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service"></param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object InternalToJson(ExchangeService service)
        {
            JsonObject jsonAttachment = base.InternalToJson(service) as JsonObject;

            jsonAttachment.Add(XmlElementNames.Item, this.item.ToJson(service, false /* isUpdateOperation */) as JsonObject);

            return jsonAttachment;
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
