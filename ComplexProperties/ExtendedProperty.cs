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
// <summary>Defines the ExtendedProperty class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an extended property.
    /// </summary>
    public sealed class ExtendedProperty : ComplexProperty
    {
        private ExtendedPropertyDefinition propertyDefinition;
        private object value;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedProperty"/> class.
        /// </summary>
        internal ExtendedProperty()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedProperty"/> class.
        /// </summary>
        /// <param name="propertyDefinition">The definition of the extended property.</param>
        internal ExtendedProperty(ExtendedPropertyDefinition propertyDefinition)
            : this()
        {
            EwsUtilities.ValidateParam(propertyDefinition, "propertyDefinition");

            this.propertyDefinition = propertyDefinition;
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.ExtendedFieldURI:
                    this.propertyDefinition = new ExtendedPropertyDefinition();
                    this.propertyDefinition.LoadFromXml(reader);
                    return true;
                case XmlElementNames.Value:
                    EwsUtilities.Assert(
                        this.PropertyDefinition != null,
                        "ExtendedProperty.TryReadElementFromXml",
                        "PropertyDefintion is missing");

                    string stringValue = reader.ReadElementValue();
                    this.value = MapiTypeConverter.ConvertToValue(this.PropertyDefinition.MapiType, stringValue);
                    return true;
                case XmlElementNames.Values:
                    EwsUtilities.Assert(
                        this.PropertyDefinition != null,
                        "ExtendedProperty.TryReadElementFromXml",
                        "PropertyDefintion is missing");

                    StringList stringList = new StringList(XmlElementNames.Value);
                    stringList.LoadFromXml(reader, reader.LocalName);
                    this.value = MapiTypeConverter.ConvertToValue(this.PropertyDefinition.MapiType, stringList);
                    return true;
                default:
                    return false;
            }
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
                    case XmlElementNames.ExtendedFieldURI:
                        this.propertyDefinition = new ExtendedPropertyDefinition();
                        this.propertyDefinition.LoadFromJson(jsonProperty.ReadAsJsonObject(key));
                        break;
                    case XmlElementNames.Value:
                        EwsUtilities.Assert(
                            this.PropertyDefinition != null,
                            "ExtendedProperty.TryReadElementFromXml",
                            "PropertyDefintion is missing");

                        string stringValue = jsonProperty.ReadAsString(key);
                        this.value = MapiTypeConverter.ConvertToValue(this.PropertyDefinition.MapiType, stringValue);
                        break;
                    case XmlElementNames.Values:
                        EwsUtilities.Assert(
                            this.PropertyDefinition != null,
                            "ExtendedProperty.TryReadElementFromXml",
                            "PropertyDefintion is missing");

                        StringList stringList = new StringList(XmlElementNames.Value);
                        ((IJsonCollectionDeserializer)stringList).CreateFromJsonCollection(jsonProperty.ReadAsArray(key), service);
                        this.value = MapiTypeConverter.ConvertToValue(this.PropertyDefinition.MapiType, stringList);
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            this.PropertyDefinition.WriteToXml(writer);

            if (MapiTypeConverter.IsArrayType(this.PropertyDefinition.MapiType))
            {
                Array array = this.Value as Array;
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Values);
                for (int index = array.GetLowerBound(0); index <= array.GetUpperBound(0); index++)
                {
                    writer.WriteElementValue(
                        XmlNamespace.Types,
                        XmlElementNames.Value,
                        MapiTypeConverter.ConvertToString(this.PropertyDefinition.MapiType, array.GetValue(index)));
                }
                writer.WriteEndElement();
            }
            else
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.Value,
                    MapiTypeConverter.ConvertToString(this.PropertyDefinition.MapiType, this.Value));
            }
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
            JsonObject jsonExtendedProp = new JsonObject();

            JsonObject jsonExtendedFieldUri = new JsonObject();
            this.PropertyDefinition.AddJsonProperties(jsonExtendedFieldUri, service);

            jsonExtendedProp.Add(XmlElementNames.ExtendedFieldURI, jsonExtendedFieldUri);

            if (MapiTypeConverter.IsArrayType(this.PropertyDefinition.MapiType))
            {
                List<object> values = new List<object>();

                foreach (object value in this.Value as Array)
                {
                    values.Add(MapiTypeConverter.ConvertToString(this.PropertyDefinition.MapiType, value));
                }

                jsonExtendedProp.Add(XmlElementNames.Values, values.ToArray());
            }
            else
            {
                jsonExtendedProp.Add(
                    XmlElementNames.Value,
                    MapiTypeConverter.ConvertToString(this.PropertyDefinition.MapiType, this.Value));
            }

            return jsonExtendedProp;
        }

        /// <summary>
        /// Gets the definition of the extended property.
        /// </summary>
        public ExtendedPropertyDefinition PropertyDefinition
        {
            get { return this.propertyDefinition; }
        }

        /// <summary>
        /// Gets or sets the value of the extended property.
        /// </summary>
        public object Value
        {
            get
            {
                return this.value;
            }

            set
            {
                EwsUtilities.ValidateParam(value, "value");
                this.SetFieldValue<object>(
                    ref this.value,
                    MapiTypeConverter.ChangeType(this.PropertyDefinition.MapiType, value));
            }
        }

        /// <summary>
        /// Gets the string value.
        /// </summary>
        /// <returns>Value as string.</returns>
        private string GetStringValue()
        {
            if (MapiTypeConverter.IsArrayType(this.PropertyDefinition.MapiType))
            {
                Array array = this.Value as Array;
                if (array == null)
                {
                    return string.Empty;
                }
                else
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append("[");
                    for (int index = array.GetLowerBound(0); index <= array.GetUpperBound(0); index++)
                    {
                        sb.Append(
                            MapiTypeConverter.ConvertToString(
                                this.PropertyDefinition.MapiType,
                                array.GetValue(index)));
                        sb.Append(",");
                    }
                    sb.Append("]");

                    return sb.ToString();
                }
            }
            else
            {
                return MapiTypeConverter.ConvertToString(this.PropertyDefinition.MapiType, this.Value);
            }
        }

        /// <summary>
        /// Determines whether the specified <see cref="T:System.Object"/> is equal to the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <param name="obj">The <see cref="T:System.Object"/> to compare with the current <see cref="T:System.Object"/>.</param>
        /// <returns>
        /// true if the specified <see cref="T:System.Object"/> is equal to the current <see cref="T:System.Object"/>; otherwise, false.
        /// </returns>
        /// <exception cref="T:System.NullReferenceException">The <paramref name="obj"/> parameter is null.</exception>
        public override bool Equals(object obj)
        {
            ExtendedProperty other = obj as ExtendedProperty;
            if ((other != null) && other.PropertyDefinition.Equals(this.PropertyDefinition))
            {
                return this.GetStringValue().Equals(other.GetStringValue());
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Serves as a hash function for a particular type.
        /// </summary>
        /// <returns>
        /// A hash code for the current <see cref="T:System.Object"/>.
        /// </returns>
        public override int GetHashCode()
        {
            return string.Concat(
                (this.PropertyDefinition != null) ? this.PropertyDefinition.GetPrintableName() : string.Empty,
                this.GetStringValue()).GetHashCode();
        }
    }
}
