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
    /// Represents an error that occurred as a result of executing a rule operation. 
    /// </summary>
    public sealed class RuleError : ComplexProperty
    {
        /// <summary>
        /// Rule property.
        /// </summary>
        private RuleProperty ruleProperty;

        /// <summary>
        /// Rule validation error code.
        /// </summary>
        private RuleErrorCode errorCode;

        /// <summary>
        /// Error message.
        /// </summary>
        private string errorMessage;

        /// <summary>
        /// Field value.
        /// </summary>
        private string value;

        /// <summary>
        /// Initializes a new instance of the <see cref="RuleError"/> class.
        /// </summary>
        internal RuleError()
            : base()
        {
        }

        /// <summary>
        /// Gets the property which failed validation.
        /// </summary>
        public RuleProperty RuleProperty
        {
            get
            {
                return this.ruleProperty;
            }
        }

        /// <summary>
        /// Gets the validation error code.
        /// </summary>
        public RuleErrorCode ErrorCode
        {
            get
            {
                return this.errorCode;
            }
        }

        /// <summary>
        /// Gets the error message.
        /// </summary>
        public string ErrorMessage
        {
            get
            {
                return this.errorMessage;
            }
        }

        /// <summary>
        /// Gets the value that failed validation.
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
                case XmlElementNames.FieldURI:
                    this.ruleProperty = reader.ReadElementValue<RuleProperty>();
                    return true;
                case XmlElementNames.ErrorCode:
                    this.errorCode = reader.ReadElementValue<RuleErrorCode>();
                    return true;
                case XmlElementNames.ErrorMessage:
                    this.errorMessage = reader.ReadElementValue();
                    return true;
                case XmlElementNames.FieldValue:
                    this.value = reader.ReadElementValue();
                    return true;
                default:
                    return false;
            }
        }
    }
}