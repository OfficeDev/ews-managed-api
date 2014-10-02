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
// <summary>Defines the ExtendedPropertyCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;

    /// <summary>
    /// Represents a collection of extended properties.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class ExtendedPropertyCollection : ComplexPropertyCollection<ExtendedProperty>, ICustomUpdateSerializer
    {
        /// <summary>
        /// Creates the complex property.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>Complex property instance.</returns>
        internal override ExtendedProperty CreateComplexProperty(string xmlElementName)
        {
            return new ExtendedProperty();
        }

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns></returns>
        internal override ExtendedProperty CreateDefaultComplexProperty()
        {
            return new ExtendedProperty();
        }

        /// <summary>
        /// Gets the name of the collection item XML element.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        /// <returns>XML element name.</returns>
        internal override string GetCollectionItemXmlElementName(ExtendedProperty complexProperty)
        {
            // This method is unused in this class, so just return null.
            return null;
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="localElementName">Name of the local element.</param>
        internal override void LoadFromXml(EwsServiceXmlReader reader, string localElementName)
        {
            ExtendedProperty extendedProperty = new ExtendedProperty();

            extendedProperty.LoadFromXml(reader, reader.LocalName);
            this.InternalAdd(extendedProperty);
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal override void WriteToXml(EwsServiceXmlWriter writer, string xmlElementName)
        {
            foreach (ExtendedProperty extendedProperty in this)
            {
                extendedProperty.WriteToXml(writer, XmlElementNames.ExtendedProperty);
            }
        }

        /// <summary>
        /// Internals to json.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns></returns>
        internal override object InternalToJson(ExchangeService service)
        {
            List<object> values = new List<object>();

            foreach (ExtendedProperty extendedProperty in this)
            {
                values.Add(extendedProperty.InternalToJson(service));
            }

            return values.ToArray();
        }

        /// <summary>
        /// Gets existing or adds new extended property.
        /// </summary>
        /// <param name="propertyDefinition">The property definition.</param>
        /// <returns>ExtendedProperty.</returns>
        private ExtendedProperty GetOrAddExtendedProperty(ExtendedPropertyDefinition propertyDefinition)
        {
            ExtendedProperty extendedProperty;
            if (!this.TryGetProperty(propertyDefinition, out extendedProperty))
            {
                extendedProperty = new ExtendedProperty(propertyDefinition);
                this.InternalAdd(extendedProperty);
            }
            return extendedProperty;
        }

        /// <summary>
        /// Sets an extended property.
        /// </summary>
        /// <param name="propertyDefinition">The property definition.</param>
        /// <param name="value">The value.</param>
        internal void SetExtendedProperty(ExtendedPropertyDefinition propertyDefinition, object value)
        {
            ExtendedProperty extendedProperty = this.GetOrAddExtendedProperty(propertyDefinition);
            extendedProperty.Value = value;
        }

        /// <summary>
        /// Removes a specific extended property definition from the collection. 
        /// </summary>
        /// <param name="propertyDefinition">The definition of the extended property to remove.</param>
        /// <returns>True if the property matching the extended property definition was successfully removed from the collection, false otherwise.</returns>
        internal bool RemoveExtendedProperty(ExtendedPropertyDefinition propertyDefinition)
        {
            EwsUtilities.ValidateParam(propertyDefinition, "propertyDefinition");

            ExtendedProperty extendedProperty;
            if (this.TryGetProperty(propertyDefinition, out extendedProperty))
            {
                return this.InternalRemove(extendedProperty);
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Tries to get property.
        /// </summary>
        /// <param name="propertyDefinition">The property definition.</param>
        /// <param name="extendedProperty">The extended property.</param>
        /// <returns>True of property exists in collection.</returns>
        private bool TryGetProperty(ExtendedPropertyDefinition propertyDefinition, out ExtendedProperty extendedProperty)
        {
            extendedProperty = this.Items.Find((prop) => prop.PropertyDefinition.Equals(propertyDefinition));
            return extendedProperty != null;
        }

        /// <summary>
        /// Tries to get property value.
        /// </summary>
        /// <param name="propertyDefinition">The property definition.</param>
        /// <param name="propertyValue">The property value.</param>
        /// <typeparam name="T">Type of expected property value.</typeparam>
        /// <returns>True if property exists in collection.</returns>
        internal bool TryGetValue<T>(ExtendedPropertyDefinition propertyDefinition, out T propertyValue)
        {
            ExtendedProperty extendedProperty;
            if (this.TryGetProperty(propertyDefinition, out extendedProperty))
            {
                // Verify that the type parameter and property definition's type are compatible.
                if (!typeof(T).IsAssignableFrom(propertyDefinition.Type))
                {
                    string errorMessage = string.Format(
                        Strings.PropertyDefinitionTypeMismatch,
                        EwsUtilities.GetPrintableTypeName(propertyDefinition.Type),
                        EwsUtilities.GetPrintableTypeName(typeof(T)));
                    throw new ArgumentException(errorMessage, "propertyDefinition");
                }

                propertyValue = (T)extendedProperty.Value;
                return true;
            }
            else
            {
                propertyValue = default(T);
                return false;
            }
        }

        #region ICustomXmlUpdateSerializer Members

        /// <summary>
        /// Writes the update to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="propertyDefinition">Property definition.</param>
        /// <returns>
        /// True if property generated serialization.
        /// </returns>
        bool ICustomUpdateSerializer.WriteSetUpdateToXml(
            EwsServiceXmlWriter writer,
            ServiceObject ewsObject,
            PropertyDefinition propertyDefinition)
        {
            List<ExtendedProperty> propertiesToSet = new List<ExtendedProperty>();

            propertiesToSet.AddRange(this.AddedItems);
            propertiesToSet.AddRange(this.ModifiedItems);

            foreach (ExtendedProperty extendedProperty in propertiesToSet)
            {
                writer.WriteStartElement(XmlNamespace.Types, ewsObject.GetSetFieldXmlElementName());
                extendedProperty.PropertyDefinition.WriteToXml(writer);

                writer.WriteStartElement(XmlNamespace.Types, ewsObject.GetXmlElementName());
                extendedProperty.WriteToXml(writer, XmlElementNames.ExtendedProperty);
                writer.WriteEndElement();

                writer.WriteEndElement();
            }

            foreach (ExtendedProperty extendedProperty in this.RemovedItems)
            {
                writer.WriteStartElement(XmlNamespace.Types, ewsObject.GetDeleteFieldXmlElementName());
                extendedProperty.PropertyDefinition.WriteToXml(writer);
                writer.WriteEndElement();
            }

            return true;
        }

        /// <summary>
        /// Writes the set update to json.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="propertyDefinition">The property definition.</param>
        /// <param name="updates">The updates.</param>
        /// <returns></returns>
        bool ICustomUpdateSerializer.WriteSetUpdateToJson(
             ExchangeService service,
             ServiceObject ewsObject,
             PropertyDefinition propertyDefinition,
             List<JsonObject> updates)
        {
            List<ExtendedProperty> propertiesToSet = new List<ExtendedProperty>();

            propertiesToSet.AddRange(this.AddedItems);
            propertiesToSet.AddRange(this.ModifiedItems);

            foreach (ExtendedProperty extendedProperty in propertiesToSet)
            {
                updates.Add(
                    PropertyBag.CreateJsonSetUpdate(
                    extendedProperty,
                    service,
                    ewsObject));
            }

            foreach (ExtendedProperty extendedProperty in this.RemovedItems)
            {
                updates.Add(PropertyBag.CreateJsonDeleteUpdate(extendedProperty.PropertyDefinition, service, ewsObject));
            }

            return true;
        }

        /// <summary>
        /// Writes the deletion update to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <returns>
        /// True if property generated serialization.
        /// </returns>
        bool ICustomUpdateSerializer.WriteDeleteUpdateToXml(EwsServiceXmlWriter writer, ServiceObject ewsObject)
        {
            foreach (ExtendedProperty extendedProperty in this.Items)
            {
                writer.WriteStartElement(XmlNamespace.Types, ewsObject.GetDeleteFieldXmlElementName());
                extendedProperty.PropertyDefinition.WriteToXml(writer);
                writer.WriteEndElement();
            }

            return true;
        }

        /// <summary>
        /// Writes the delete update to json.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="updates">The updates.</param>
        /// <returns></returns>
        bool ICustomUpdateSerializer.WriteDeleteUpdateToJson(ExchangeService service, ServiceObject ewsObject, List<JsonObject> updates)
        {
            foreach (ExtendedProperty extendedProperty in this.Items)
            {
                updates.Add(PropertyBag.CreateJsonDeleteUpdate(extendedProperty.PropertyDefinition, service, ewsObject));
            }

            return true;
        }

        #endregion
    }
}
