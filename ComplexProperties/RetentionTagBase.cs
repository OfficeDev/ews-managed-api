// ---------------------------------------------------------------------------
// <copyright file="RetentionTagBase.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the RetentionTagBase class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Text;

    /// <summary>
    /// Represents the retention tag of an item.
    /// </summary>
    public class RetentionTagBase : ComplexProperty
    {
        /// <summary>
        /// Xml element name.
        /// </summary>
        private readonly string xmlElementName;

        /// <summary>
        /// Is explicit.
        /// </summary>
        private bool isExplicit;

        /// <summary>
        /// Retention id.
        /// </summary>
        private Guid retentionId;

        /// <summary>
        /// Initializes a new instance of the <see cref="RetentionTagBase"/> class.
        /// </summary>
        /// <param name="xmlElementName">Xml element name.</param>
        public RetentionTagBase(string xmlElementName)
        {
            this.xmlElementName = xmlElementName;
        }

        /// <summary>
        /// Gets or sets if the tag is explicit.
        /// </summary>
        public bool IsExplicit
        {
            get { return this.isExplicit; }
            set { this.SetFieldValue<bool>(ref this.isExplicit, value); }
        }

        /// <summary>
        /// Gets or sets the retention id.
        /// </summary>
        public Guid RetentionId
        {
            get { return this.retentionId; }
            set { this.SetFieldValue<Guid>(ref this.retentionId, value); }
        }

        /// <summary>
        /// Reads attributes from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
        {
            this.isExplicit = reader.ReadAttributeValue<bool>(XmlAttributeNames.IsExplicit);
        }

        /// <summary>
        /// Reads text value from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadTextValueFromXml(EwsServiceXmlReader reader)
        {
            this.retentionId = new Guid(reader.ReadValue());
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
                    case XmlAttributeNames.IsExplicit:
                        this.isExplicit = jsonProperty.ReadAsBool(key);
                        break;
                    case JsonObject.JsonValueString:
                        this.retentionId = new Guid(jsonProperty.ReadAsString(key));
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
            writer.WriteAttributeValue(XmlAttributeNames.IsExplicit, this.isExplicit);
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (this.retentionId != null && this.retentionId != Guid.Empty)
            {
                writer.WriteValue(this.retentionId.ToString(), this.xmlElementName);
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

            jsonProperty.Add(XmlAttributeNames.IsExplicit, this.isExplicit);

            if (this.retentionId != null && this.retentionId != Guid.Empty)
            {
                jsonProperty.Add(JsonObject.JsonValueString, this.retentionId);
            }

            return jsonProperty;
        }

        #region Object method overrides

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override string ToString()
        {
            if (this.retentionId == null || this.retentionId == Guid.Empty)
            {
                return string.Empty;
            }
            else
            {
                return this.retentionId.ToString();
            }
        }

        #endregion
    }
}
