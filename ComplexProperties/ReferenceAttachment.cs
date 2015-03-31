namespace Microsoft.Exchange.WebServices.Data
{
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

    /// <summary>
    /// Represents a Reference attachment.
    /// </summary>
    public sealed class ReferenceAttachment : Attachment
    {
        private String attachLongPathName;
        private String providerType;
        private Int32 permissionType;

        /// <summary>
        /// Initializes a new instance of the <see cref="ReferenceAttachment"/> class.
        /// </summary>
        /// <param name="owner">The owner.</param>
        internal ReferenceAttachment(Item owner)
            : base(owner)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ReferenceAttachment"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal ReferenceAttachment(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.ReferenceAttachment;
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        /// <param name="attachmentIndex">Index of this attachment.</param>
        internal override void Validate(int attachmentIndex)
        {

        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            bool result = base.TryReadElementFromXml(reader);

            if (!result)
            {
                switch (reader.LocalName)
                {
                    case XmlElementNames.AttachLongPathName:
                        this.attachLongPathName = reader.ReadElementValue();
                        break;
                    case XmlElementNames.ProviderType:
                        this.providerType  = reader.ReadElementValue();
                        break;
                    case XmlElementNames.PermissionType:
                        this.permissionType = reader.ReadElementValue<Int32>();
                        break;
                }

            }

            return result;
        }

        /// <summary>
        /// For ReferenceAttachment, the only thing need to patch is the AttachmentId.
        /// </summary>
        /// <param name="reader"></param>
        /// <returns></returns>
        internal override bool TryReadElementFromXmlToPatch(EwsServiceXmlReader reader)
        {
            return base.TryReadElementFromXml(reader);
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service"></param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            base.LoadFromJson(jsonProperty, service);

            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlElementNames.AttachLongPathName:
                        this.attachLongPathName = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.ProviderType:
                        this.providerType = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.PermissionType:
                        this.PermissionType = jsonProperty.ReadAsInt(key);
                        break;  
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Writes elements and content to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            base.WriteElementsToXml(writer);
            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Content);
            writer.WriteEndElement();
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

       


            return jsonAttachment;
        }





        /// <summary>
        /// Gets the value of the AttachLongPathName of the referance Attachment
        /// </summary>
        public string AttachLongPathName
        {
            get { return this.attachLongPathName; }
            set { //this.attachLongPathName = value; 
            }
        }
        /// <summary>
        /// Gets the value of the ProviderType of the referance Attachment
        /// </summary>
        public string ProviderType
        {
            get { return this.providerType; }
            set { //this.providerType = value; 
            }
        }
        /// <summary>
        /// Gets the value of the PermissionType of the referance Attachment
        /// </summary>
        public Int32 PermissionType
        {
            get { return this.permissionType; }
            set { //this.permissionType = value; 
            }
        }
    
    }
}

