#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

//-----------------------------------------------------------------------
// <summary>Defines the AlternateIdBase class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents the base class for Id expressed in a specific format.
    /// </summary>
    public abstract class AlternateIdBase : ISelfValidate, IJsonSerializable
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AlternateIdBase"/> class.
        /// </summary>
        internal AlternateIdBase()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AlternateIdBase"/> class.
        /// </summary>
        /// <param name="format">The format.</param>
        internal AlternateIdBase(IdFormat format)
            : this()
        {
            this.Format = format;
        }

        /// <summary>
        /// Gets or sets the format in which the Id in expressed.
        /// </summary>
        public IdFormat Format
        {
            get; set;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal abstract string GetXmlElementName();

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal virtual void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteAttributeValue(XmlAttributeNames.Format, this.Format);
        }

        /// <summary>
        /// Loads the attributes from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal virtual void LoadAttributesFromXml(EwsServiceXmlReader reader)
        {
            this.Format = reader.ReadAttributeValue<IdFormat>(XmlAttributeNames.Format);
        }

        /// <summary>
        /// Loads the attributes from json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        internal virtual void LoadAttributesFromJson(JsonObject responseObject)
        {
            this.Format = responseObject.ReadEnumValue<IdFormat>(XmlAttributeNames.Format);
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal void WriteToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Types, this.GetXmlElementName());

            this.WriteAttributesToXml(writer);

            writer.WriteEndElement(); // this.GetXmlElementName()
        }

        /// <summary>
        /// Creates a JSON representation of this object.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        object IJsonSerializable.ToJson(ExchangeService service)
        {
            JsonObject jsonObject = new JsonObject();

            this.InternalToJson(jsonObject);

            return jsonObject;
        }

        /// <summary>
        /// Creates a JSON representation of this object..
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        internal virtual void InternalToJson(JsonObject jsonObject)
        {
            jsonObject.Add(XmlAttributeNames.Format, this.Format);
            jsonObject.AddTypeParameter(this.GetXmlElementName());
        }

        /// <summary>
        /// Validate this instance.
        /// </summary>
        internal virtual void InternalValidate()
        {
            // nothing to do.
        }

        #region ISelfValidate Members

        /// <summary>
        /// Validates this instance.
        /// </summary>
        void ISelfValidate.Validate()
        {
            this.InternalValidate();
        }

        #endregion
    }
}
