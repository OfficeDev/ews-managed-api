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
    using System.IO;

    /// <summary>
    /// Represents an ContactPhoneEntity object.
    /// </summary>
    public sealed class ContactPhoneEntity : ComplexProperty
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ContactPhoneEntity"/> class.
        /// </summary>
        internal ContactPhoneEntity()
            : base()
        {
        }

        /// <summary>
        /// Gets the phone entity OriginalPhoneString.
        /// </summary>
        public string OriginalPhoneString { get; internal set; }

        /// <summary>
        /// Gets the phone entity PhoneString.
        /// </summary>
        public string PhoneString { get; internal set; }

        /// <summary>
        /// Gets the phone entity Type.
        /// </summary>
        public string Type { get; internal set; }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.NlgOriginalPhoneString:
                    this.OriginalPhoneString = reader.ReadElementValue();
                    return true;

                case XmlElementNames.NlgPhoneString:
                    this.PhoneString = reader.ReadElementValue();
                    return true;

                case XmlElementNames.NlgType:
                    this.Type = reader.ReadElementValue();
                    return true;

                default:
                    return base.TryReadElementFromXml(reader);
            }
        }
    }
}