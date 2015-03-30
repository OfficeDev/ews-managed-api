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
    /// <summary>
    /// Represents an AQS highlight term. 
    /// </summary>
    public sealed class HighlightTerm : ComplexProperty
    {
        /// <summary>
        /// Term scope.
        /// </summary>
        private string scope;

        /// <summary>
        /// Term value.
        /// </summary>
        private string value;

        /// <summary>
        /// Initializes a new instance of the <see cref="HighlightTerm"/> class.
        /// </summary>
        internal HighlightTerm()
            : base()
        {
        }

        /// <summary>
        /// Gets term scope.
        /// </summary>
        public string Scope
        {
            get
            {
                return this.scope;
            }
        }

        /// <summary>
        /// Gets term value.
        /// </summary>
        public string Value
        {
            get
            {
                return this.value;
            }
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.HighlightTermScope:
                    this.scope = reader.ReadElementValue();
                    return true;
                case XmlElementNames.HighlightTermValue:
                    this.value = reader.ReadElementValue();
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            if (jsonProperty.ContainsKey(XmlElementNames.HighlightTermScope))
            {
                this.scope = jsonProperty.ReadAsString(XmlElementNames.HighlightTermScope);
            }

            if (jsonProperty.ContainsKey(XmlElementNames.HighlightTermValue))
            {
                this.value = jsonProperty.ReadAsString(XmlElementNames.HighlightTermValue);
            }
        }
    }
}