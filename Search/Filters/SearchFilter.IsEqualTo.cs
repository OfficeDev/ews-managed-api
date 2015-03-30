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