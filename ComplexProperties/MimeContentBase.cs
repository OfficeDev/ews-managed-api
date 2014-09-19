// ---------------------------------------------------------------------------
// <copyright file="MimeContentBase.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MimeContentBase class.</summary>
//-----------------------------------------------------------------------

using System.Security.Cryptography;

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Text;

    /// <summary>
    /// Represents the MIME content of an item.
    /// </summary>
    public abstract class MimeContentBase : ComplexProperty
    {
        /// <summary>
        /// characterSet returned 
        /// </summary>
        private string characterSet;

        /// <summary>
        /// content received
        /// </summary>
        private byte[] content;
    
        /// <summary>
        /// Reads attributes from XML.
        /// This should always be UTF-8 for MimeContentUTF8
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
        {
            this.characterSet = reader.ReadAttributeValue<string>(XmlAttributeNames.CharacterSet);
        }

        /// <summary>
        /// Reads text value from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadTextValueFromXml(EwsServiceXmlReader reader)
        {
            this.content = System.Convert.FromBase64String(reader.ReadValue());
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlAttributeNames.CharacterSet:
                        this.characterSet = jsonProperty.ReadAsString(key);
                        break;
                    case JsonObject.JsonValueString:
                        this.content = jsonProperty.ReadAsBase64Content(key);
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Writes attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteAttributeValue(XmlAttributeNames.CharacterSet, this.CharacterSet);
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (this.Content != null && this.Content.Length > 0)
            {
                writer.WriteBase64ElementValue(this.Content);
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

            jsonProperty.Add(XmlAttributeNames.ChangeKey, this.CharacterSet);

            if (this.Content != null && this.Content.Length > 0)
            {
                jsonProperty.AddBase64(JsonObject.JsonValueString, this.Content);
            }

            return jsonProperty;
        }

        /// <summary>
        /// Gets or sets the character set of the content.
        /// </summary>
        public string CharacterSet
        {
            get { return this.characterSet; }
            set { this.SetFieldValue<string>(ref this.characterSet, value); }
        }

        /// <summary>
        /// Gets or sets the content.
        /// </summary>
        public byte[] Content
        {
            get { return this.content; }
            set { this.SetFieldValue<byte[]>(ref this.content, value); }
        }
    }
}
