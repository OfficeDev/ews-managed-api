// ---------------------------------------------------------------------------
// <copyright file="SearchFilter.IsEqualTo.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the IsEqualTo class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <content>
    /// Contains nested type SearchFilter.IsEqualTo.
    /// </content>
    public abstract partial class SearchFilter
    {
        /// <summary>
        /// Represents a search filter that checks if a property is equal to a given value or other property.
        /// </summary>
        public sealed class IsEqualTo : RelationalFilter
        {
            /// <summary>
            /// Initializes a new instance of the <see cref="IsEqualTo"/> class.
            /// </summary>
            public IsEqualTo()
                : base()
            {
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="IsEqualTo"/> class.
            /// </summary>
            /// <param name="propertyDefinition">The definition of the property that is being compared. Property definitions are available as static members from schema classes (for example, EmailMessageSchema.Subject, AppointmentSchema.Start, ContactSchema.GivenName, etc.)</param>
            /// <param name="otherPropertyDefinition">The definition of the property to compare with. Property definitions are available on schema classes (EmailMessageSchema, AppointmentSchema, etc.)</param>
            public IsEqualTo(PropertyDefinitionBase propertyDefinition, PropertyDefinitionBase otherPropertyDefinition)
                : base(propertyDefinition, otherPropertyDefinition)
            {
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="IsEqualTo"/> class.
            /// </summary>
            /// <param name="propertyDefinition">The definition of the property that is being compared. Property definitions are available as static members from schema classes (for example, EmailMessageSchema.Subject, AppointmentSchema.Start, ContactSchema.GivenName, etc.)</param>
            /// <param name="value">The value to compare the property with.</param>
            public IsEqualTo(PropertyDefinitionBase propertyDefinition, object value)
                : base(propertyDefinition, value)
            {
            }

            /// <summary>
            /// Gets the name of the XML element.
            /// </summary>
            /// <returns>XML element name.</returns>
            internal override string GetXmlElementName()
            {
                return XmlElementNames.IsEqualTo;
            }
        }
    }
}