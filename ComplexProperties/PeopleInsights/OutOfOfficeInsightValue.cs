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
    using System.Globalization;
    using System.Xml;

    /// <summary>
    /// Represents the OutOfOfficeInsightValue.
    /// </summary>
    public sealed class OutOfOfficeInsightValue : InsightValue
    {
        private DateTime startTime;
        private DateTime endTime;
        private string culture = CultureInfo.CurrentCulture.Name;
        private string message;

        /// <summary>
        /// Get the start date and time.
        /// </summary>
        public DateTime StartTime
        {
            get
            {
                return this.startTime;
            }
        }

        /// <summary>
        /// Get the end date and time.
        /// </summary>
        public DateTime EndTime
        {
            get
            { 
                return this.endTime; 
            }
        }

        /// <summary>
        /// Get the culture of the reply.
        /// </summary>
        public string Culture
        {
            get 
            { 
                return this.culture; 
            }
        }

        /// <summary>
        /// Get the reply message.
        /// </summary>
        public string Message
        {
            get
            { 
                return this.message; 
            }
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">XML reader</param>
        /// <returns>Whether the element was read</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.InsightSource:
                    this.InsightSource = reader.ReadElementValue<string>();
                    break;
                case XmlElementNames.StartTime:
                    this.startTime = reader.ReadElementValueAsDateTime(XmlNamespace.Types, XmlElementNames.StartTime).Value;
                    break;
                case XmlElementNames.EndTime:
                    this.endTime = reader.ReadElementValueAsDateTime(XmlNamespace.Types, XmlElementNames.EndTime).Value;
                    break;
                case XmlElementNames.Culture:
                    this.culture = reader.ReadElementValue();
                    break;
                case XmlElementNames.Message:
                    this.message = reader.ReadElementValue();
                    break;
                default:
                    return false;
            }

            return true;
        }
    }
}
