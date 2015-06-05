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
            /// Writes the elements to XML.
            /// </summary>
            /// <param name="writer">The writer.</param>
            internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
            {
                this.PropertyDefinition.WriteToXml(writer);
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