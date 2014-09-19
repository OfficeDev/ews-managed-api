// ---------------------------------------------------------------------------
// <copyright file="SearchFilter.RelationalFilter.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the RelationalFilter class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;

    /// <content>
    /// Contains nested type SearchFilter.RelationalFilter.
    /// </content>
    public abstract partial class SearchFilter
    {
        /// <summary>
        /// Represents the base class for relational filters (for example, IsEqualTo, IsGreaterThan or IsLessThanOrEqualTo).
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never)]
        public abstract class RelationalFilter : PropertyBasedFilter
        {
            private PropertyDefinitionBase otherPropertyDefinition;
            private object value;

            /// <summary>
            /// Initializes a new instance of the <see cref="RelationalFilter"/> class.
            /// </summary>
            internal RelationalFilter()
                : base()
            {
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="RelationalFilter"/> class.
            /// </summary>
            /// <param name="propertyDefinition">The definition of the property that is being compared. Property definitions are available as static members from schema classes (for example, EmailMessageSchema.Subject, AppointmentSchema.Start, ContactSchema.GivenName, etc.)</param>
            /// <param name="otherPropertyDefinition">The definition of the property to compare with. Property definitions are available as static members from schema classes (for example, EmailMessageSchema, AppointmentSchema, etc.)</param>
            internal RelationalFilter(PropertyDefinitionBase propertyDefinition, PropertyDefinitionBase otherPropertyDefinition)
                : base(propertyDefinition)
            {
                this.otherPropertyDefinition = otherPropertyDefinition;
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="RelationalFilter"/> class.
            /// </summary>
            /// <param name="propertyDefinition">The definition of the property that is being compared. Property definitions are available as static members from schema classes (for example, EmailMessageSchema.Subject, AppointmentSchema.Start, ContactSchema.GivenName, etc.)</param>
            /// <param name="value">The value to compare with.</param>
            internal RelationalFilter(PropertyDefinitionBase propertyDefinition, object value)
                : base(propertyDefinition)
            {
                this.value = value;
            }

            /// <summary>
            /// Validate instance.
            /// </summary>
            internal override void InternalValidate()
            {
                base.InternalValidate();

                if (this.otherPropertyDefinition == null && this.value == null)
                {
                    throw new ServiceValidationException(Strings.EqualityComparisonFilterIsInvalid);
                }
                else if (value != null)
                {
                    // All common value types (String, Int32, DateTime, ...) implement IConvertible.
                    // Value types that don't implement IConvertible must implement ISearchStringProvider 
                    // in order to be used in a search filter.
                    if (!((value is IConvertible) || (value is ISearchStringProvider)))
                    {
                        throw new ServiceValidationException(
                            string.Format(Strings.SearchFilterComparisonValueTypeIsNotSupported, value.GetType().Name));
                    }
                }
            }

            /// <summary>
            /// Tries to read element from XML.
            /// </summary>
            /// <param name="reader">The reader.</param>
            /// <returns>True if element was read.</returns>
            internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
            {
                bool result = base.TryReadElementFromXml(reader);

                if (!result)
                {
                    if (reader.LocalName == XmlElementNames.FieldURIOrConstant)
                    {
                        reader.Read();
                        reader.EnsureCurrentNodeIsStartElement();

                        if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Constant))
                        {
                            this.value = reader.ReadAttributeValue(XmlAttributeNames.Value);

                            result = true;
                        }
                        else
                        {
                            result = PropertyDefinitionBase.TryLoadFromXml(reader, ref this.otherPropertyDefinition);
                        }
                    }
                }

                return result;
            }

            /// <summary>
            /// Loads from json.
            /// </summary>
            /// <param name="jsonProperty">The json property.</param>
            /// <param name="service">The service.</param>
            internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
            {
                base.LoadFromJson(jsonProperty, service);

                JsonObject jsonFieldUriOrConstant = jsonProperty.ReadAsJsonObject(XmlElementNames.FieldURIOrConstant).ReadAsJsonObject(XmlElementNames.Item);

                if (jsonFieldUriOrConstant.ReadTypeString() == XmlElementNames.Constant)
                {
                    this.value = jsonFieldUriOrConstant[XmlElementNames.Value];
                }
                else
                {
                    this.otherPropertyDefinition = PropertyDefinitionBase.TryLoadFromJson(jsonProperty);
                }
            }

            /// <summary>
            /// Writes the elements to XML.
            /// </summary>
            /// <param name="writer">The writer.</param>
            internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
            {
                base.WriteElementsToXml(writer);

                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.FieldURIOrConstant);

                if (this.Value != null)
                {
                    writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Constant);
                    writer.WriteAttributeValue(XmlAttributeNames.Value, true /* alwaysWriteEmptyString */, this.Value);
                    writer.WriteEndElement(); // Constant
                }
                else
                {
                    this.OtherPropertyDefinition.WriteToXml(writer);
                }

                writer.WriteEndElement(); // FieldURIOrConstant
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
                JsonObject jsonFilter = base.InternalToJson(service) as JsonObject;
                JsonObject jsonFieldUriOrConstant = new JsonObject();

                if (this.Value != null)
                {
                    JsonObject jsonConstant = new JsonObject();
                    jsonConstant.Add(XmlElementNames.Value, this.Value);
                    jsonConstant.AddTypeParameter(XmlElementNames.Constant);
                    jsonFieldUriOrConstant.Add(XmlElementNames.Item, jsonConstant);
                }
                else
                {
                    jsonFieldUriOrConstant.Add(XmlElementNames.Item, ((IJsonSerializable)this.OtherPropertyDefinition).ToJson(service));
                }

                jsonFilter.Add(XmlElementNames.FieldURIOrConstant, jsonFieldUriOrConstant);

                return jsonFilter;
            }

            /// <summary>
            /// Gets or sets the definition of the property to compare with. Property definitions are available as static members
            /// from schema classes (for example, EmailMessageSchema.Subject, AppointmentSchema.Start, ContactSchema.GivenName, etc.)
            /// The OtherPropertyDefinition and Value properties are mutually exclusive; setting one resets the other to null.
            /// </summary>
            public PropertyDefinitionBase OtherPropertyDefinition
            {
                get
                {
                    return this.otherPropertyDefinition;
                }

                set
                {
                    this.SetFieldValue<PropertyDefinitionBase>(ref this.otherPropertyDefinition, value);
                    this.value = null;
                }
            }

            /// <summary>
            /// Gets or sets the value to compare with. The Value and OtherPropertyDefinition properties
            /// are mutually exclusive; setting one resets the other to null.
            /// </summary>
            public object Value
            {
                get
                {
                    return this.value;
                }

                set
                {
                    this.SetFieldValue<object>(ref this.value, value);
                    this.otherPropertyDefinition = null;
                }
            }
        }
    }
}