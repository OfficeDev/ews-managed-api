// ---------------------------------------------------------------------------
// <copyright file="Attachment.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the Attachment class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an attachment to an item.
    /// </summary>
    public abstract class Attachment : ComplexProperty
    {
        private Item owner;
        private string id;
        private string name;
        private string contentType;
        private string contentId;
        private string contentLocation;
        private int size;
        private DateTime lastModifiedTime;
        private bool isInline;
        private ExchangeService service;

        /// <summary>
        /// Initializes a new instance of the <see cref="Attachment"/> class.
        /// </summary>
        /// <param name="owner">The owner.</param>
        internal Attachment(Item owner)
        {
            this.owner = owner;

            if (owner != null)
            {
                this.service = this.owner.Service;
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Attachment"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal Attachment(ExchangeService service)
        {
            this.service = service;
        }

        /// <summary>
        /// Throws exception if this is not a new service object.
        /// </summary>
        internal void ThrowIfThisIsNotNew()
        {
            if (!this.IsNew)
            {
                throw new InvalidOperationException(Strings.AttachmentCannotBeUpdated);
            }
        }

        /// <summary>
        /// Sets value of field.
        /// </summary>
        /// <remarks>
        /// We override the base implementation. Attachments cannot be modified so any attempts
        /// the change a property on an existing attachment is an error.
        /// </remarks>
        /// <typeparam name="T">Field type.</typeparam>
        /// <param name="field">The field.</param>
        /// <param name="value">The value.</param>
        internal override void SetFieldValue<T>(ref T field, T value)
        {
            this.ThrowIfThisIsNotNew();
            base.SetFieldValue<T>(ref field, value);
        }

        /// <summary>
        /// Gets the Id of the attachment.
        /// </summary>
        public string Id
        {
            get { return this.id; }
            internal set { this.id = value; }
        }

        /// <summary>
        /// Gets or sets the name of the attachment.
        /// </summary>
        public string Name
        {
            get { return this.name; }
            set { this.SetFieldValue<string>(ref this.name, value); }
        }

        /// <summary>
        /// Gets or sets the content type of the attachment.
        /// </summary>
        public string ContentType
        {
            get { return this.contentType; }
            set { this.SetFieldValue<string>(ref this.contentType, value); }
        }

        /// <summary>
        /// Gets or sets the content Id of the attachment. ContentId can be used as a custom way to identify
        /// an attachment in order to reference it from within the body of the item the attachment belongs to.
        /// </summary>
        public string ContentId
        {
            get { return this.contentId; }
            set { this.SetFieldValue<string>(ref this.contentId, value); }
        }

        /// <summary>
        /// Gets or sets the content location of the attachment. ContentLocation can be used to associate
        /// an attachment with a Url defining its location on the Web.
        /// </summary>
        public string ContentLocation
        {
            get { return this.contentLocation; }
            set { this.SetFieldValue<string>(ref this.contentLocation, value); }
        }

        /// <summary>
        /// Gets the size of the attachment.
        /// </summary>
        public int Size
        {
            get
            {
                EwsUtilities.ValidatePropertyVersion(this.service, ExchangeVersion.Exchange2010, "Size");

                return this.size;
            }

            internal set
            {
                EwsUtilities.ValidatePropertyVersion(this.service, ExchangeVersion.Exchange2010, "Size");

                this.SetFieldValue<int>(ref this.size, value);
            }
        }

        /// <summary>
        /// Gets the date and time when this attachment was last modified.
        /// </summary>
        public DateTime LastModifiedTime
        {
            get
            {
                EwsUtilities.ValidatePropertyVersion(this.service, ExchangeVersion.Exchange2010, "LastModifiedTime");

                return this.lastModifiedTime;
            }

            internal set
            {
                EwsUtilities.ValidatePropertyVersion(this.service, ExchangeVersion.Exchange2010, "LastModifiedTime");

                this.SetFieldValue<DateTime>(ref this.lastModifiedTime, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this is an inline attachment.
        /// Inline attachments are not visible to end users.
        /// </summary>
        public bool IsInline
        {
            get
            {
                EwsUtilities.ValidatePropertyVersion(this.service, ExchangeVersion.Exchange2010, "IsInline");

                return this.isInline;
            }

            set
            {
                EwsUtilities.ValidatePropertyVersion(this.service, ExchangeVersion.Exchange2010, "IsInline");

                this.SetFieldValue<bool>(ref this.isInline, value);
            }
        }

        /// <summary>
        /// True if the attachment has not yet been saved, false otherwise.
        /// </summary>
        internal bool IsNew
        {
            get { return string.IsNullOrEmpty(this.Id); }
        }

        /// <summary>
        /// Gets the owner of the attachment.
        /// </summary>
        internal Item Owner
        {
            get { return this.owner; }
        }

        /// <summary>
        /// Gets the related exchange service.
        /// </summary>
        internal ExchangeService Service
        {
            get { return this.service; }
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal abstract string GetXmlElementName();

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.AttachmentId:
                    this.id = reader.ReadAttributeValue(XmlAttributeNames.Id);

                    if (this.Owner != null)
                    {
                        string rootItemChangeKey = reader.ReadAttributeValue(XmlAttributeNames.RootItemChangeKey);

                        if (!string.IsNullOrEmpty(rootItemChangeKey))
                        {
                            this.Owner.RootItemId.ChangeKey = rootItemChangeKey;
                        }
                    }
                    reader.ReadEndElementIfNecessary(XmlNamespace.Types, XmlElementNames.AttachmentId);
                    return true;
                case XmlElementNames.Name:
                    this.name = reader.ReadElementValue();
                    return true;
                case XmlElementNames.ContentType:
                    this.contentType = reader.ReadElementValue();
                    return true;
                case XmlElementNames.ContentId:
                    this.contentId = reader.ReadElementValue();
                    return true;
                case XmlElementNames.ContentLocation:
                    this.contentLocation = reader.ReadElementValue();
                    return true;
                case XmlElementNames.Size:
                    this.size = reader.ReadElementValue<int>();
                    return true;
                case XmlElementNames.LastModifiedTime:
                    this.lastModifiedTime = reader.ReadElementValueAsDateTime().Value;
                    return true;
                case XmlElementNames.IsInline:
                    this.isInline = reader.ReadElementValue<bool>();
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service"></param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlElementNames.AttachmentId:
                        this.LoadAttachmentIdFromJson(jsonProperty.ReadAsJsonObject(key));
                        break;
                    case XmlElementNames.Name:
                        this.name = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.ContentType:
                        this.contentType = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.ContentId:
                        this.contentId = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.ContentLocation:
                        this.contentLocation = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.Size:
                        this.size = jsonProperty.ReadAsInt(key);
                        break;
                    case XmlElementNames.LastModifiedTime:
                        this.lastModifiedTime = service.ConvertUniversalDateTimeStringToLocalDateTime(jsonProperty.ReadAsString(key)).Value;
                        break;
                    case XmlElementNames.IsInline:
                        this.isInline = jsonProperty.ReadAsBool(key);
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Loads the attachment id from json.
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        private void LoadAttachmentIdFromJson(JsonObject jsonObject)
        {
            this.id = jsonObject.ReadAsString(XmlAttributeNames.Id);

            if (this.Owner != null &&
                jsonObject.ContainsKey(XmlAttributeNames.RootItemChangeKey))
            {
                string rootItemChangeKey = jsonObject.ReadAsString(XmlAttributeNames.RootItemChangeKey);

                if (!string.IsNullOrEmpty(rootItemChangeKey))
                {
                    this.Owner.RootItemId.ChangeKey = rootItemChangeKey;
                }
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Name, this.Name);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.ContentType, this.ContentType);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.ContentId, this.ContentId);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.ContentLocation, this.ContentLocation);
            if (writer.Service.RequestedServerVersion > ExchangeVersion.Exchange2007_SP1)
            {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsInline, this.IsInline);
            }
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
            JsonObject jsonProperty = new JsonObject();

            jsonProperty.AddTypeParameter(this.GetXmlElementName());
            jsonProperty.Add(XmlElementNames.Name, this.Name);
            jsonProperty.Add(XmlElementNames.ContentType, this.ContentType);
            jsonProperty.Add(XmlElementNames.ContentId, this.ContentId);
            jsonProperty.Add(XmlElementNames.ContentLocation, this.ContentLocation);
            if (service.RequestedServerVersion > ExchangeVersion.Exchange2007_SP1)
            {
                jsonProperty.Add(XmlElementNames.IsInline, this.IsInline);
            }

            return jsonProperty;
        }

        /// <summary>
        /// Load the attachment.
        /// </summary>
        /// <param name="bodyType">Type of the body.</param>
        /// <param name="additionalProperties">The additional properties.</param>
        internal void InternalLoad(BodyType? bodyType, IEnumerable<PropertyDefinitionBase> additionalProperties)
        {
            this.service.GetAttachment(
                this,
                bodyType,
                additionalProperties);
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        /// <param name="attachmentIndex">Index of this attachment.</param>
        internal virtual void Validate(int attachmentIndex)
        {
        }

        /// <summary>
        /// Loads the attachment. Calling this method results in a call to EWS.
        /// </summary>
        public void Load()
        {
            this.InternalLoad(null, null);
        }
    }
}
