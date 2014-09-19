// ---------------------------------------------------------------------------
// <copyright file="OofReply.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the OofReply class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Text;

    /// <summary>
    /// Represents an Out of Office response.
    /// </summary>
    public sealed class OofReply
    {
        private string culture = CultureInfo.CurrentCulture.Name;
        private string message;

        /// <summary>
        /// Writes an empty OofReply to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal static void WriteEmptyReplyToXml(EwsServiceXmlWriter writer, string xmlElementName)
        {
            writer.WriteStartElement(XmlNamespace.Types, xmlElementName);
            writer.WriteEndElement(); // xmlElementName
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="OofReply"/> class.
        /// </summary>
        public OofReply()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="OofReply"/> class.
        /// </summary>
        /// <param name="message">The reply message.</param>
        public OofReply(string message)
        {
            this.message = message;
        }

        /// <summary>
        /// Defines an implicit conversion between string an OofReply.
        /// </summary>
        /// <param name="message">The message to convert into OofReply.</param>
        /// <returns>An OofReply initialized with the specified message.</returns>
        public static implicit operator OofReply(string message)
        {
            return new OofReply(message);
        }

        /// <summary>
        /// Defines an implicit conversion between OofReply and string.
        /// </summary>
        /// <param name="oofReply">The OofReply to convert into a string.</param>
        /// <returns>A string containing the message of the specified OofReply.</returns>
        public static implicit operator string(OofReply oofReply)
        {
            EwsUtilities.ValidateParam(oofReply, "oofReply");

            return oofReply.Message;
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal void LoadFromXml(EwsServiceXmlReader reader, string xmlElementName)
        {
            reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, xmlElementName);

            if (reader.HasAttributes)
            {
                this.culture = reader.ReadAttributeValue("xml:lang");
            }

            this.message = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Message);

            reader.ReadEndElement(XmlNamespace.Types, xmlElementName);
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        /// <param name="service">The service.</param>
        internal void LoadFromJson(JsonObject jsonObject, ExchangeService service)
        {
            if (jsonObject.ContainsKey("xml:lang"))
            {
                this.culture = jsonObject.ReadAsString("xml:lang");
            }
            this.message = jsonObject.ReadAsString(XmlElementNames.Message);
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal void WriteToXml(EwsServiceXmlWriter writer, string xmlElementName)
        {
            writer.WriteStartElement(XmlNamespace.Types, xmlElementName);

            if (this.Culture != null)
            {
                writer.WriteAttributeValue(
                    "xml",
                    "lang",
                    this.Culture);
            }

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.Message,
                this.Message);

            writer.WriteEndElement(); // xmlElementName
        }

        /// <summary>
        /// Serializes to json.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns></returns>
        internal JsonObject InternalToJson(ExchangeService service)
        {
            JsonObject jsonProperty = new JsonObject();

            if (this.Culture != null)
            {
                jsonProperty.Add(
                    "xml:lang",
                    this.Culture);
            }

            jsonProperty.Add(XmlElementNames.Message, this.Message);

            return jsonProperty;
        }

        /// <summary>
        /// Obtains a string representation of the reply.
        /// </summary>
        /// <returns>A string containing the reply message.</returns>
        public override string ToString()
        {
            return this.Message;
        }

        /// <summary>
        /// Gets or sets the culture of the reply.
        /// </summary>
        public string Culture
        {
            get { return this.culture; }
            set { this.culture = value; }
        }

        /// <summary>
        /// Gets or sets the reply message.
        /// </summary>
        public string Message
        {
            get { return this.message; }
            set { this.message = value; }
        }
    }
}
