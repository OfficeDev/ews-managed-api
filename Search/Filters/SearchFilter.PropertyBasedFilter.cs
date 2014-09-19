// ---------------------------------------------------------------------------
// <copyright file="SearchFilter.PropertyBasedFilter.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PropertyBasedFilter class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;

    /// <content>
    /// Contains nested type SearchFilter.PropertyBasedFilter.
    /// </content>
    public abstract partial class SearchFilter
    {
        /// <summary>
        /// Represents a search filter where an item or folder property is involved.
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never)]
        public abstract class PropertyBasedFilter : SearchFilter
        {
            private PropertyDefinitionBase propertyDefinition;

            /// <summary>
            /// Initializes a new instance of the <see cref="PropertyBasedFilter"/> class.
            /// </summary>
            internal PropertyBasedFilter()
                : base()
            {
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="PropertyBasedFilter"/> class.
            /// </summary>
            /// <param name="propertyDefinition">The property definition.</param>
            internal PropertyBasedFilter(PropertyDefinitionBase propertyDefinition)
                : base()
            {
                this.propertyDefinition = propertyDefinition;
            }

            /// <summary>
            /// Validate instance.
            /// </summary>
            internal override void InternalValidate()
            {
                if (this.propertyDefinition == null)
                {
                    throw new ServiceValidationException(Strings.PropertyDefinitionPropertyMustBeSet);
                }
            }

            /// <summary>
            /// Tries to read element from XML.
            /// </summary>
            /// <param name="reader">The reader.</param>
            /// <returns>True if element was read.</returns>
            internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
            {
                return PropertyDefinitionBase.TryLoadFromXml(reader, ref this.propertyDefinition);
            }

            /// <summary>
            /// Loads from json.
            /// </summary>
            /// <param name="jsonProperty">The json property.</param>
            /// <param name="service">The service.</param>
            internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
            {
                this.PropertyDefinition = PropertyDefinitionBase.TryLoadFromJson(jsonProperty.ReadAsJsonObject(XmlElementNames.Item));
            }

            /// <summary>
            /// Writes the elements to XML.
            /// </summary>
            /// <param name="writer">The writer.</param>
            internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
            {
                this.PropertyDefinition.WriteToXml(writer);
            }

            internal override object InternalToJson(ExchangeService service)
            {
                JsonObject jsonFilter = base.InternalToJson(service) as JsonObject;

                jsonFilter.Add(XmlElementNames.Item, ((IJsonSerializable)this.PropertyDefinition).ToJson(service));

                return jsonFilter;
            }

            /// <summary>
            /// Gets or sets the definition of the property that is involved in the search filter. Property definitions are
            /// available as static members from schema classes (for example, EmailMessageSchema.Subject, AppointmentSchema.Start, ContactSchema.GivenName, etc.)
            /// </summary>
            public PropertyDefinitionBase PropertyDefinition
            {
                get { return this.propertyDefinition; }
                set { this.SetFieldValue<PropertyDefinitionBase>(ref this.propertyDefinition, value); }
            }
        }
    }
}