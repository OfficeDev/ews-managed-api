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
    using System.Text;

    /// <summary>
    /// Represents the definition of an extended property.
    /// </summary>
    public sealed class ExtendedPropertyDefinition : PropertyDefinitionBase
    {
        #region Constants

        private const string FieldFormat = "{0}: {1} ";

        private const string PropertySetFieldName = "PropertySet";
        private const string PropertySetIdFieldName = "PropertySetId";
        private const string TagFieldName = "Tag";
        private const string NameFieldName = "Name";
        private const string IdFieldName = "Id";
        private const string MapiTypeFieldName = "MapiType";

        #endregion

        #region Fields

        private DefaultExtendedPropertySet? propertySet;
        private Guid? propertySetId;
        private int? tag;
        private string name;
        private int? id;
        private MapiPropertyType mapiType;

        #endregion

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedPropertyDefinition"/> class.
        /// </summary>
        internal ExtendedPropertyDefinition()
            : base()
        {
            this.mapiType = MapiPropertyType.String;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedPropertyDefinition"/> class.
        /// </summary>
        /// <param name="mapiType">The MAPI type of the extended property.</param>
        internal ExtendedPropertyDefinition(MapiPropertyType mapiType)
            : this()
        {
            this.mapiType = mapiType;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedPropertyDefinition"/> class.
        /// </summary>
        /// <param name="tag">The tag of the extended property.</param>
        /// <param name="mapiType">The MAPI type of the extended property.</param>
        public ExtendedPropertyDefinition(int tag, MapiPropertyType mapiType)
            : this(mapiType)
        {
            if (tag < 0 || tag > UInt16.MaxValue)
            {
                throw new ArgumentOutOfRangeException("tag", Strings.TagValueIsOutOfRange);
            }
            this.tag = tag;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedPropertyDefinition"/> class.
        /// </summary>
        /// <param name="propertySet">The extended property set of the extended property.</param>
        /// <param name="name">The name of the extended property.</param>
        /// <param name="mapiType">The MAPI type of the extended property.</param>
        public ExtendedPropertyDefinition(
            DefaultExtendedPropertySet propertySet,
            string name,
            MapiPropertyType mapiType)
            : this(mapiType)
        {
            EwsUtilities.ValidateParam(name, "name");

            this.propertySet = propertySet;
            this.name = name;
        }

        /// <summary>
        /// Initializes a new instance of ExtendedPropertyDefinition.
        /// </summary>
        /// <param name="propertySet">The property set of the extended property.</param>
        /// <param name="id">The Id of the extended property.</param>
        /// <param name="mapiType">The MAPI type of the extended property.</param>
        public ExtendedPropertyDefinition(
            DefaultExtendedPropertySet propertySet,
            int id,
            MapiPropertyType mapiType)
            : this(mapiType)
        {
            this.propertySet = propertySet;
            this.id = id;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedPropertyDefinition"/> class.
        /// </summary>
        /// <param name="propertySetId">The property set Id of the extended property.</param>
        /// <param name="name">The name of the extended property.</param>
        /// <param name="mapiType">The MAPI type of the extended property.</param>
        public ExtendedPropertyDefinition(
            Guid propertySetId,
            string name,
            MapiPropertyType mapiType)
            : this(mapiType)
        {
            EwsUtilities.ValidateParam(name, "name");

            this.propertySetId = propertySetId;
            this.name = name;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedPropertyDefinition"/> class.
        /// </summary>
        /// <param name="propertySetId">The property set Id of the extended property.</param>
        /// <param name="id">The Id of the extended property.</param>
        /// <param name="mapiType">The MAPI type of the extended property.</param>
        public ExtendedPropertyDefinition(
            Guid propertySetId,
            int id,
            MapiPropertyType mapiType)
            : this(mapiType)
        {
            this.propertySetId = propertySetId;
            this.id = id;
        }

        /// <summary>
        /// Determines whether two specified instances of ExtendedPropertyDefinition are equal.
        /// </summary>
        /// <param name="extPropDef1">First extended property definition.</param>
        /// <param name="extPropDef2">Second extended property definition.</param>
        /// <returns>True if extended property definitions are equal.</returns>
        internal static bool IsEqualTo(ExtendedPropertyDefinition extPropDef1, ExtendedPropertyDefinition extPropDef2)
        {
            return
                object.ReferenceEquals(extPropDef1, extPropDef2) ||
                ((object)extPropDef1 != null &&
                 (object)extPropDef2 != null &&
                 extPropDef1.Id == extPropDef2.Id &&
                 extPropDef1.MapiType == extPropDef2.MapiType &&
                 extPropDef1.Tag == extPropDef2.Tag &&
                 extPropDef1.Name == extPropDef2.Name &&
                 extPropDef1.PropertySet == extPropDef2.PropertySet &&
                 extPropDef1.propertySetId == extPropDef2.propertySetId);
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.ExtendedFieldURI;
        }

        /// <summary>
        /// Gets the type for json.
        /// </summary>
        /// <returns></returns>
        protected override string GetJsonType()
        {
            return JsonNames.PathToExtendedFieldType;
        }

        /// <summary>
        /// Gets the minimum Exchange version that supports this extended property.
        /// </summary>
        /// <value>The version.</value>
        public override ExchangeVersion Version
        {
            get { return ExchangeVersion.Exchange2007_SP1; }
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            if (this.propertySet.HasValue)
            {
                writer.WriteAttributeValue(XmlAttributeNames.DistinguishedPropertySetId, this.propertySet.Value);
            }
            if (this.propertySetId.HasValue)
            {
                writer.WriteAttributeValue(XmlAttributeNames.PropertySetId, this.propertySetId.Value.ToString());
            }
            if (this.tag.HasValue)
            {
                writer.WriteAttributeValue(XmlAttributeNames.PropertyTag, this.tag.Value);
            }
            if (!string.IsNullOrEmpty(this.name))
            {
                writer.WriteAttributeValue(XmlAttributeNames.PropertyName, this.name);
            }
            if (this.id.HasValue)
            {
                writer.WriteAttributeValue(XmlAttributeNames.PropertyId, this.id.Value);
            }
            writer.WriteAttributeValue(XmlAttributeNames.PropertyType, this.mapiType);
        }

        /// <summary>
        /// Adds the json properties.
        /// </summary>
        /// <param name="jsonPropertyDefinition">The json property definition.</param>
        /// <param name="service">The service.</param>
        internal override void AddJsonProperties(JsonObject jsonPropertyDefinition, ExchangeService service)
        {
            if (this.propertySet.HasValue)
            {
                jsonPropertyDefinition.Add(XmlAttributeNames.DistinguishedPropertySetId, this.propertySet.Value);
            }
            if (this.propertySetId.HasValue)
            {
                jsonPropertyDefinition.Add(XmlAttributeNames.PropertySetId, this.propertySetId.Value.ToString());
            }
            if (this.tag.HasValue)
            {
                jsonPropertyDefinition.Add(XmlAttributeNames.PropertyTag, this.tag.Value);
            }
            if (!string.IsNullOrEmpty(this.name))
            {
                jsonPropertyDefinition.Add(XmlAttributeNames.PropertyName, this.name);
            }
            if (this.id.HasValue)
            {
                jsonPropertyDefinition.Add(XmlAttributeNames.PropertyId, this.id.Value);
            }
            jsonPropertyDefinition.Add(XmlAttributeNames.PropertyType, this.mapiType);
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadFromXml(EwsServiceXmlReader reader)
        {
            string attributeValue;

            attributeValue = reader.ReadAttributeValue(XmlAttributeNames.DistinguishedPropertySetId);
            if (!string.IsNullOrEmpty(attributeValue))
            {
                this.propertySet = (DefaultExtendedPropertySet)Enum.Parse(typeof(DefaultExtendedPropertySet), attributeValue, false);
            }

            attributeValue = reader.ReadAttributeValue(XmlAttributeNames.PropertySetId);
            if (!string.IsNullOrEmpty(attributeValue))
            {
                this.propertySetId = new Guid(attributeValue);
            }

            attributeValue = reader.ReadAttributeValue(XmlAttributeNames.PropertyTag);
            if (!string.IsNullOrEmpty(attributeValue))
            {
                this.tag = Convert.ToUInt16(attributeValue, 16);
            }

            this.name = reader.ReadAttributeValue(XmlAttributeNames.PropertyName);

            attributeValue = reader.ReadAttributeValue(XmlAttributeNames.PropertyId);
            if (!string.IsNullOrEmpty(attributeValue))
            {
                this.id = int.Parse(attributeValue);
            }

            this.mapiType = reader.ReadAttributeValue<MapiPropertyType>(XmlAttributeNames.PropertyType);
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        internal void LoadFromJson(JsonObject jsonObject)
        {
            foreach (string key in jsonObject.Keys)
            {
                switch (key)
                {
                    case XmlAttributeNames.DistinguishedPropertySetId:
                        this.propertySet = jsonObject.ReadEnumValue<DefaultExtendedPropertySet>(key);
                        break;
                    case XmlAttributeNames.PropertySetId:
                        this.propertySetId = new Guid(jsonObject.ReadAsString(key));
                        break;
                    case XmlAttributeNames.PropertyTag:
                        this.tag = Convert.ToUInt16(jsonObject.ReadAsString(key), 16);
                        break;
                    case XmlAttributeNames.PropertyName:
                        this.name = jsonObject.ReadAsString(key);
                        break;
                    case XmlAttributeNames.PropertyId:
                        this.id = jsonObject.ReadAsInt(key);
                        break;
                    case XmlAttributeNames.PropertyType:
                        this.mapiType = jsonObject.ReadEnumValue<MapiPropertyType>(key);
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Determines whether two specified instances of ExtendedPropertyDefinition are equal.
        /// </summary>
        /// <param name="extPropDef1">First extended property definition.</param>
        /// <param name="extPropDef2">Second extended property definition.</param>
        /// <returns>True if extended property definitions are equal.</returns>
        public static bool operator ==(ExtendedPropertyDefinition extPropDef1, ExtendedPropertyDefinition extPropDef2)
        {
            return ExtendedPropertyDefinition.IsEqualTo(extPropDef1, extPropDef2);
        }

        /// <summary>
        /// Determines whether two specified instances of ExtendedPropertyDefinition are not equal.
        /// </summary>
        /// <param name="extPropDef1">First extended property definition.</param>
        /// <param name="extPropDef2">Second extended property definition.</param>
        /// <returns>True if extended property definitions are equal.</returns>
        public static bool operator !=(ExtendedPropertyDefinition extPropDef1, ExtendedPropertyDefinition extPropDef2)
        {
            return !ExtendedPropertyDefinition.IsEqualTo(extPropDef1, extPropDef2);
        }

        /// <summary>
        /// Determines whether a given extended property definition is equal to this extended property definition.
        /// </summary>
        /// <param name="obj">The object to check for equality.</param>
        /// <returns>True if the properties definitions define the same extended property.</returns>
        public override bool Equals(object obj)
        {
            ExtendedPropertyDefinition propertyDefinition = obj as ExtendedPropertyDefinition;
            return ExtendedPropertyDefinition.IsEqualTo(propertyDefinition, this);
        }

        /// <summary>
        /// Serves as a hash function for a particular type.
        /// </summary>
        /// <returns>
        /// A hash code for the current <see cref="T:System.Object"/>.
        /// </returns>
        public override int GetHashCode()
        {
            return this.GetPrintableName().GetHashCode();
        }

        /// <summary>
        /// Gets the property definition's printable name.
        /// </summary>
        /// <returns>
        /// The property definition's printable name.
        /// </returns>
        internal override string GetPrintableName()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("{");
            sb.Append(FormatField<string>(NameFieldName, this.Name));
            sb.Append(FormatField<MapiPropertyType?>(MapiTypeFieldName, this.MapiType));
            sb.Append(FormatField<int?>(IdFieldName, this.Id));
            sb.Append(FormatField<DefaultExtendedPropertySet?>(PropertySetFieldName, this.PropertySet));
            sb.Append(FormatField<Guid?>(PropertySetIdFieldName, this.PropertySetId));
            sb.Append(FormatField<int?>(TagFieldName, this.Tag));
            sb.Append("}");
            return sb.ToString();
        }

        /// <summary>
        /// Formats the field.
        /// </summary>
        /// <typeparam name="T">Type of field value.</typeparam>
        /// <param name="name">The name.</param>
        /// <param name="fieldValue">The field value.</param>
        /// <returns>Formatted value.</returns>
        internal string FormatField<T>(string name, T fieldValue)
        {
            return (fieldValue != null)
                        ? string.Format(FieldFormat, name, fieldValue.ToString())
                        : string.Empty;
        }

        /// <summary>
        /// Gets the property set of the extended property.
        /// </summary>
        public DefaultExtendedPropertySet? PropertySet
        {
            get { return this.propertySet; }
        }

        /// <summary>
        /// Gets the property set Id or the extended property.
        /// </summary>
        public Guid? PropertySetId
        {
            get { return this.propertySetId; }
        }

        /// <summary>
        /// Gets the extended property's tag.
        /// </summary>
        public int? Tag
        {
            get { return this.tag; }
        }

        /// <summary>
        /// Gets the name of the extended property.
        /// </summary>
        public string Name
        {
            get { return this.name; }
        }

        /// <summary>
        /// Gets the Id of the extended property.
        /// </summary>
        public int? Id
        {
            get { return this.id; }
        }

        /// <summary>
        /// Gets the MAPI type of the extended property.
        /// </summary>
        public MapiPropertyType MapiType
        {
            get { return this.mapiType; }
        }

        /// <summary>
        /// Gets the property type.
        /// </summary>
        public override Type Type
        {
            get { return MapiTypeConverter.MapiTypeConverterMap[this.MapiType].Type; }
        }
    }
}