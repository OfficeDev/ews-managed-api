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
    /// Contains nested type Recurrence.ContainsSubstring.
    /// </content>
    public abstract partial class SearchFilter
    {
        /// <summary>
        /// Represents a search filter that checks for the presence of a substring inside a text property.
        /// Applications can use ContainsSubstring to define conditions such as "Field CONTAINS Value" or "Field IS PREFIXED WITH Value".
        /// </summary>
        public sealed class ContainsSubstring : PropertyBasedFilter
        {
            private ContainmentMode containmentMode = ContainmentMode.Substring;
            private ComparisonMode comparisonMode = ComparisonMode.IgnoreCase;
            private string value;

            /// <summary>
            /// Initializes a new instance of the <see cref="ContainsSubstring"/> class.
            /// </summary>
            public ContainsSubstring()
                : base()
            {
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="ContainsSubstring"/> class.
            /// The ContainmentMode property is initialized to ContainmentMode.Substring, and 
            /// the ComparisonMode property is initialized to ComparisonMode.IgnoreCase.
            /// </summary>
            /// <param name="propertyDefinition">The definition of the property that is being compared. Property definitions are available as static members from schema classes (for example, EmailMessageSchema.Subject, AppointmentSchema.Start, ContactSchema.GivenName, etc.)</param>
            /// <param name="value">The value to compare with.</param>
            public ContainsSubstring(PropertyDefinitionBase propertyDefinition, string value)
                : base(propertyDefinition)
            {
                this.value = value;
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="ContainsSubstring"/> class.
            /// </summary>
            /// <param name="propertyDefinition">The definition of the property that is being compared. Property definitions are available as static members from schema classes (for example, EmailMessageSchema.Subject, AppointmentSchema.Start, ContactSchema.GivenName, etc.)</param>
            /// <param name="value">The value to compare with.</param>
            /// <param name="containmentMode">The containment mode.</param>
            /// <param name="comparisonMode">The comparison mode.</param>
            public ContainsSubstring(
                PropertyDefinitionBase propertyDefinition,
                string value,
                ContainmentMode containmentMode,
                ComparisonMode comparisonMode)
                : this(propertyDefinition, value)
            {
                this.containmentMode = containmentMode;
                this.comparisonMode = comparisonMode;
            }

            /// <summary>
            /// Validate instance.
            /// </summary>
            internal override void InternalValidate()
            {
                base.InternalValidate();

                if (string.IsNullOrEmpty(this.value))
                {
                    throw new ServiceValidationException(Strings.ValuePropertyMustBeSet);
                }
            }

            /// <summary>
            /// Gets the name of the XML element.
            /// </summary>
            /// <returns>XML element name.</returns>
            internal override string GetXmlElementName()
            {
                return XmlElementNames.Contains;
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
                    if (reader.LocalName == XmlElementNames.Constant)
                    {
                        this.value = reader.ReadAttributeValue(XmlAttributeNames.Value);

                        result = true;
                    }
                }

                return result;
            }

            /// <summary>
            /// Reads the attributes from XML.
            /// </summary>
            /// <param name="reader">The reader.</param>
            internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
            {
                base.ReadAttributesFromXml(reader);

                this.containmentMode = reader.ReadAttributeValue<ContainmentMode>(XmlAttributeNames.ContainmentMode);

                try
                {
                    this.comparisonMode = reader.ReadAttributeValue<ComparisonMode>(XmlAttributeNames.ContainmentComparison);
                }
                catch (ArgumentException)
                {
                    // This will happen if we receive a value that is defined in the EWS schema but that is not defined
                    // in the API (see the comments in ComparisonMode.cs). We map that value to IgnoreCaseAndNonSpacingCharacters.
                    this.comparisonMode = ComparisonMode.IgnoreCaseAndNonSpacingCharacters;
                }
            }

            /// <summary>
            /// Writes the attributes to XML.
            /// </summary>
            /// <param name="writer">The writer.</param>
            internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
            {
                base.WriteAttributesToXml(writer);

                writer.WriteAttributeValue(XmlAttributeNames.ContainmentMode, this.ContainmentMode);
                writer.WriteAttributeValue(XmlAttributeNames.ContainmentComparison, this.ComparisonMode);
            }

            /// <summary>
            /// Writes the elements to XML.
            /// </summary>
            /// <param name="writer">The writer.</param>
            internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
            {
                base.WriteElementsToXml(writer);

                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Constant);
                writer.WriteAttributeValue(XmlAttributeNames.Value, this.Value);
                writer.WriteEndElement(); // Constant
            }

            /// <summary>
            /// Gets or sets the containment mode.
            /// </summary>
            public ContainmentMode ContainmentMode
            {
                get { return this.containmentMode; }
                set { this.SetFieldValue<ContainmentMode>(ref this.containmentMode, value); }
            }

            /// <summary>
            /// Gets or sets the comparison mode.
            /// </summary>
            public ComparisonMode ComparisonMode
            {
                get { return this.comparisonMode; }
                set { this.SetFieldValue<ComparisonMode>(ref this.comparisonMode, value); }
            }

            /// <summary>
            /// Gets or sets the value to compare the specified property with.
            /// </summary>
            public string Value
            {
                get { return this.value; }
                set { this.SetFieldValue<string>(ref this.value, value); }
            }
        }
    }
}