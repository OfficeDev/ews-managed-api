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
// <summary>Defines the PropertyDefinitionBase class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents the base class for all property definitions.
    /// </summary>
    [Serializable]
    public abstract class PropertyDefinitionBase : IJsonSerializable
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PropertyDefinitionBase"/> class.
        /// </summary>
        internal PropertyDefinitionBase()
        {
        }

        /// <summary>
        /// Tries to load from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="propertyDefinition">The property definition.</param>
        /// <returns>True if property was loaded.</returns>
        internal static bool TryLoadFromXml(EwsServiceXmlReader reader, ref PropertyDefinitionBase propertyDefinition)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.FieldURI:
                    propertyDefinition = ServiceObjectSchema.FindPropertyDefinition(reader.ReadAttributeValue(XmlAttributeNames.FieldURI));
                    reader.SkipCurrentElement();
                    return true;
                case XmlElementNames.IndexedFieldURI:
                    propertyDefinition = new IndexedPropertyDefinition(
                        reader.ReadAttributeValue(XmlAttributeNames.FieldURI),
                        reader.ReadAttributeValue(XmlAttributeNames.FieldIndex));
                    reader.SkipCurrentElement();
                    return true;
                case XmlElementNames.ExtendedFieldURI:
                    propertyDefinition = new ExtendedPropertyDefinition();
                    (propertyDefinition as ExtendedPropertyDefinition).LoadFromXml(reader);
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Tries to load from XML.
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        /// <returns>True if property was loaded.</returns>
        internal static PropertyDefinitionBase TryLoadFromJson(JsonObject jsonObject)
        {
            switch (jsonObject.ReadTypeString())
            {
                case JsonNames.PathToUnindexedFieldType:
                    return ServiceObjectSchema.FindPropertyDefinition(jsonObject.ReadAsString(XmlAttributeNames.FieldURI));
                case JsonNames.PathToIndexedFieldType:
                    return new IndexedPropertyDefinition(
                        jsonObject.ReadAsString(XmlAttributeNames.FieldURI),
                        jsonObject.ReadAsString(XmlAttributeNames.FieldIndex));
                case JsonNames.PathToExtendedFieldType:
                    ExtendedPropertyDefinition propertyDefinition = new ExtendedPropertyDefinition();
                    propertyDefinition.LoadFromJson(jsonObject);
                    return propertyDefinition;
                default:
                    return null;
            }
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal abstract string GetXmlElementName();

        /// <summary>
        /// Gets the type for json.
        /// </summary>
        /// <returns></returns>
        protected abstract string GetJsonType();

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal abstract void WriteAttributesToXml(EwsServiceXmlWriter writer);

        /// <summary>
        /// Gets the minimum Exchange version that supports this property.
        /// </summary>
        /// <value>The version.</value>
        public abstract ExchangeVersion Version { get; }

        /// <summary>
        /// Gets the property definition's printable name.
        /// </summary>
        /// <returns>The property definition's printable name.</returns>
        internal abstract string GetPrintableName();

        /// <summary>
        /// Gets the type of the property.
        /// </summary>
        public abstract Type Type { get; }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal virtual void WriteToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Types, this.GetXmlElementName());
            this.WriteAttributesToXml(writer);
            writer.WriteEndElement();
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
            JsonObject jsonPropertyDefinition = new JsonObject();

            jsonPropertyDefinition.AddTypeParameter(this.GetJsonType());
            this.AddJsonProperties(jsonPropertyDefinition, service);

            return jsonPropertyDefinition;
        }

        /// <summary>
        /// Adds the json properties.
        /// </summary>
        /// <param name="jsonPropertyDefinition">The json property definition.</param>
        /// <param name="service">The service.</param>
        internal abstract void AddJsonProperties(JsonObject jsonPropertyDefinition, ExchangeService service);

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override string ToString()
        {
            return this.GetPrintableName();
        }
    }
}