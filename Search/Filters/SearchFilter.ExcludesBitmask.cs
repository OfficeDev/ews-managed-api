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
// <summary>Defines the ExcludesBitmask class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <content>
    /// Contains nested type SearchFilter.ExcludesBitmask.
    /// </content>
    public abstract partial class SearchFilter
    {
        /// <summary>
        /// Represents a bitmask exclusion search filter. Applications can use ExcludesBitExcludesBitmaskFilter to define
        /// conditions such as "(OrdinalField and 0x0010) != 0x0010"
        /// </summary>
        public sealed class ExcludesBitmask : PropertyBasedFilter
        {
            private int bitmask;

            /// <summary>
            /// Initializes a new instance of the <see cref="ExcludesBitmask"/> class.
            /// </summary>
            public ExcludesBitmask()
                : base()
            {
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="ExcludesBitmask"/> class.
            /// </summary>
            /// <param name="propertyDefinition">The definition of the property that is being compared. Property definitions are available as static members from schema classes (for example, EmailMessageSchema.Subject, AppointmentSchema.Start, ContactSchema.GivenName, etc.)</param>
            /// <param name="bitmask">The bitmask to compare with.</param>
            public ExcludesBitmask(PropertyDefinitionBase propertyDefinition, int bitmask)
                : base(propertyDefinition)
            {
                this.bitmask = bitmask;
            }

            /// <summary>
            /// Gets the name of the XML element.
            /// </summary>
            /// <returns>XML element name.</returns>
            internal override string GetXmlElementName()
            {
                return XmlElementNames.Excludes;
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
                    if (reader.LocalName == XmlElementNames.Bitmask)
                    {
                        // EWS always returns the Bitmask value in hexadecimal
                        this.bitmask = Convert.ToInt32(reader.ReadAttributeValue(XmlAttributeNames.Value), 16);
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

                this.bitmask = Convert.ToInt32(jsonProperty.ReadAsJsonObject(XmlElementNames.Bitmask).ReadAsString(XmlElementNames.Value), 16);
            }

            /// <summary>
            /// Writes the elements to XML.
            /// </summary>
            /// <param name="writer">The writer.</param>
            internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
            {
                base.WriteElementsToXml(writer);

                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Bitmask);
                writer.WriteAttributeValue(XmlAttributeNames.Value, this.Bitmask);
                writer.WriteEndElement(); // Bitmask
            }

            /// <summary>
            /// Internals to json.
            /// </summary>
            /// <param name="service">The service.</param>
            /// <returns></returns>
            internal override object InternalToJson(ExchangeService service)
            {
                JsonObject jsonfilter = base.InternalToJson(service) as JsonObject;

                JsonObject jsonBitmask = new JsonObject();
                jsonBitmask.Add(XmlElementNames.Value, "0x" + this.Bitmask.ToString("X", System.Globalization.CultureInfo.InvariantCulture));

                return jsonfilter;
            }

            /// <summary>
            /// Gets or sets the bitmask to compare the property with.
            /// </summary>
            public int Bitmask
            {
                get { return this.bitmask; }
                set { this.SetFieldValue<int>(ref this.bitmask, value); }
            }
        }
    }
}