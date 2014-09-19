// ---------------------------------------------------------------------------
// <copyright file="PhoneCallId.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PhoneCallId class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the Id of a phone call.
    /// </summary>
    internal sealed class PhoneCallId : ComplexProperty
    {
        private string id;

        /// <summary>
        /// Initializes a new instance of the <see cref="PhoneCallId"/> class.
        /// </summary>
        internal PhoneCallId()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PhoneCallId"/> class. 
        /// </summary>
        /// <param name="id">The Id of the phone call.</param>
        internal PhoneCallId(string id)
        {
            this.id = id;
        }

        /// <summary>
        /// Reads attributes from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
        {
            this.id = reader.ReadAttributeValue(XmlAttributeNames.Id);
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            this.id = jsonProperty.ReadAsString(XmlAttributeNames.Id);
        }

        /// <summary>
        /// Writes attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteAttributeValue(XmlAttributeNames.Id, this.id);
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal void WriteToXml(EwsServiceXmlWriter writer)
        {
            this.WriteToXml(writer, XmlElementNames.PhoneCallId);
        }

        /// <summary>
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object InternalToJson(ExchangeService service)
        {
            JsonObject jsonProperty = new JsonObject();

            jsonProperty.Add(XmlAttributeNames.Id, this.id);

            return jsonProperty;
        }

        /// <summary>
        /// Gets or sets the Id of the phone call.
        /// </summary>
        internal string Id
        {
            get
            {
                return this.id;
            }

            set
            {
                this.id = value;
            }
        }
    }
}
