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
    /// Represents abstract complex property definition.
    /// </summary>
    internal abstract class ComplexPropertyDefinitionBase : PropertyDefinition
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ComplexPropertyDefinitionBase"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        internal ComplexPropertyDefinitionBase(
            string xmlElementName,
            PropertyDefinitionFlags flags,
            ExchangeVersion version)
            : base(
                xmlElementName,
                flags,
                version)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ComplexPropertyDefinitionBase"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="version">The version.</param>
        internal ComplexPropertyDefinitionBase(
            string xmlElementName,
            string uri,
            ExchangeVersion version)
            : base(
                xmlElementName,
                uri,
                version)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ComplexPropertyDefinitionBase"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        internal ComplexPropertyDefinitionBase(
            string xmlElementName,
            string uri,
            PropertyDefinitionFlags flags,
            ExchangeVersion version)
            : base(
                xmlElementName,
                uri,
                flags,
                version)
        {
        }

        /// <summary>
        /// Creates the property instance.
        /// </summary>
        /// <param name="owner">The owner.</param>
        /// <returns>ComplexProperty.</returns>
        internal abstract ComplexProperty CreatePropertyInstance(ServiceObject owner);

        /// <summary>
        /// Internals the load from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="propertyBag">The property bag.</param>
        internal virtual void InternalLoadFromXml(EwsServiceXmlReader reader, PropertyBag propertyBag)
        {
            object complexProperty;

            bool justCreated = GetPropertyInstance(propertyBag, out complexProperty);

            if (!justCreated && this.HasFlag(PropertyDefinitionFlags.UpdateCollectionItems, propertyBag.Owner.Service.RequestedServerVersion))
            {
                (complexProperty as ComplexProperty).UpdateFromXml(reader, reader.LocalName);
            }
            else
            {
                (complexProperty as ComplexProperty).LoadFromXml(reader, reader.LocalName);
            }

            propertyBag[this] = complexProperty;
        }

        /// <summary>
        /// Internals the load from json.
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        /// <param name="service">The service.</param>
        /// <param name="propertyBag">The property bag.</param>
        internal virtual void InternalLoadFromJson(JsonObject jsonObject, ExchangeService service, PropertyBag propertyBag)
        {
            object complexProperty;

            bool justCreated = GetPropertyInstance(propertyBag, out complexProperty);

            (complexProperty as ComplexProperty).LoadFromJson(jsonObject, service);

            propertyBag[this] = complexProperty;
        }

        /// <summary>
        /// Internals the load colelction from json.
        /// </summary>
        /// <param name="jsonCollection">The json collection.</param>
        /// <param name="service">The service.</param>
        /// <param name="propertyBag">The property bag.</param>
        private void InternalLoadCollectionFromJson(object[] jsonCollection, ExchangeService service, PropertyBag propertyBag)
        {
            object propertyInstance;

            bool justCreated = GetPropertyInstance(propertyBag, out propertyInstance);

            IJsonCollectionDeserializer complexProperty = propertyInstance as IJsonCollectionDeserializer;

            if (complexProperty == null)
            {
                throw new ServiceJsonDeserializationException();
            }

            if (!justCreated && this.HasFlag(PropertyDefinitionFlags.UpdateCollectionItems, propertyBag.Owner.Service.RequestedServerVersion))
            {
                complexProperty.UpdateFromJsonCollection(jsonCollection, service);
            }
            else
            {
                complexProperty.CreateFromJsonCollection(jsonCollection, service);
            }

            propertyBag[this] = complexProperty;
        }

        /// <summary>
        /// Gets the property instance.
        /// </summary>
        /// <param name="propertyBag">The property bag.</param>
        /// <param name="complexProperty">The property instance.</param>
        /// <returns>True if the instance is newly created.</returns>
        private bool GetPropertyInstance(PropertyBag propertyBag, out object complexProperty)
        {
            complexProperty = null;
            if (!propertyBag.TryGetValue(this, out complexProperty) || !this.HasFlag(PropertyDefinitionFlags.ReuseInstance, propertyBag.Owner.Service.RequestedServerVersion))
            {
                complexProperty = this.CreatePropertyInstance(propertyBag.Owner);
                return true;
            }

            return false;
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="propertyBag">The property bag.</param>
        internal override sealed void LoadPropertyValueFromXml(EwsServiceXmlReader reader, PropertyBag propertyBag)
        {
            reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, this.XmlElementName);

            if (!reader.IsEmptyElement || reader.HasAttributes)
            {
                this.InternalLoadFromXml(reader, propertyBag);
            }

            reader.ReadEndElementIfNecessary(XmlNamespace.Types, this.XmlElementName);
        }

        /// <summary>
        /// Loads the property value from json.
        /// </summary>
        /// <param name="value">The JSON value.  Can be a JsonObject, string, number, bool, array, or null.</param>
        /// <param name="service">The service.</param>
        /// <param name="propertyBag">The property bag.</param>
        internal override void LoadPropertyValueFromJson(object value, ExchangeService service, PropertyBag propertyBag)
        {
            JsonObject jsonObject = value as JsonObject;

            if (jsonObject != null)
            {
                this.InternalLoadFromJson(jsonObject, service, propertyBag);
            }
            else if (value.GetType().IsArray)
            {
                this.InternalLoadCollectionFromJson(value as object[], service, propertyBag);
            }
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="propertyBag">The property bag.</param>
        /// <param name="isUpdateOperation">Indicates whether the context is an update operation.</param>
        internal override void WritePropertyValueToXml(
            EwsServiceXmlWriter writer,
            PropertyBag propertyBag,
            bool isUpdateOperation)
        {
            ComplexProperty complexProperty = (ComplexProperty)propertyBag[this];

            if (complexProperty != null)
            {
                complexProperty.WriteToXml(writer, this.XmlElementName);
            }
        }

        /// <summary>
        /// Writes the json value.
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        /// <param name="propertyBag">The property bag.</param>
        /// <param name="service">The service.</param>
        /// <param name="isUpdateOperation">if set to <c>true</c> [is update operation].</param>
        internal override void WriteJsonValue(JsonObject jsonObject, PropertyBag propertyBag, ExchangeService service, bool isUpdateOperation)
        {
            ComplexProperty complexProperty = (ComplexProperty)propertyBag[this];

            if (complexProperty != null)
            {
                jsonObject.Add(this.XmlElementName, complexProperty.InternalToJson(service));
            }
        }
    }
}