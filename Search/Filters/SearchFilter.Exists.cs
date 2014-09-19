// ---------------------------------------------------------------------------
// <copyright file="SearchFilter.Exists.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the Exists class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <content>
    /// Contains nested type SearchFilter.Exists.
    /// </content>
    public abstract partial class SearchFilter
    {
        /// <summary>
        /// Represents a search filter checking if a field is set. Applications can use
        /// ExistsFilter to define conditions such as "Field IS SET".
        /// </summary>
        public sealed class Exists : PropertyBasedFilter
        {
            /// <summary>
            /// Initializes a new instance of the <see cref="Exists"/> class.
            /// </summary>
            public Exists()
                : base()
            {
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="Exists"/> class.
            /// </summary>
            /// <param name="propertyDefinition">The definition of the property to check the existence of. Property definitions are available as static members from schema classes (for example, EmailMessageSchema.Subject, AppointmentSchema.Start, ContactSchema.GivenName, etc.)</param>
            public Exists(PropertyDefinitionBase propertyDefinition)
                : base(propertyDefinition)
            {
            }

            /// <summary>
            /// Gets the name of the XML element.
            /// </summary>
            /// <returns>XML element name.</returns>
            internal override string GetXmlElementName()
            {
                return XmlElementNames.Exists;
            }
        }
    }
}