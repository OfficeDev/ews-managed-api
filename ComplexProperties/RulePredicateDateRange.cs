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

    /// <summary>
    /// Represents the date and time range within which messages have been received.
    /// </summary>
    public sealed class RulePredicateDateRange : ComplexProperty
    {
        /// <summary>
        /// The start DateTime.
        /// </summary>
        private DateTime? start;

        /// <summary>
        /// The end DateTime.
        /// </summary>
        private DateTime? end;

        /// <summary>
        /// Initializes a new instance of the <see cref="RulePredicateDateRange"/> class.
        /// </summary>
        internal RulePredicateDateRange()
            : base()
        {
        }

        /// <summary>
        /// Gets or sets the range start date and time. If Start is set to null, no 
        /// start date applies.
        /// </summary>
        public DateTime? Start
        {
            get
            {
                return this.start;
            }

            set
            {
                this.SetFieldValue<DateTime?>(ref this.start, value);
            }
        }

        /// <summary>
        /// Gets or sets the range end date and time. If End is set to null, no end 
        /// date applies.
        /// </summary>
        public DateTime? End
        {
            get
            {
                return this.end;
            }

            set
            {
                this.SetFieldValue<DateTime?>(ref this.end, value);
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
                case XmlElementNames.StartDateTime:
                    this.start = reader.ReadElementValueAsDateTime();
                    return true;
                case XmlElementNames.EndDateTime:
                    this.end = reader.ReadElementValueAsDateTime();
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (this.Start.HasValue)
            {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.StartDateTime, this.Start.Value);
            }
            if (this.End.HasValue)
            {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.EndDateTime, this.End.Value);
            }
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void InternalValidate()
        {
            base.InternalValidate();
            if (this.start.HasValue &&
                this.end.HasValue &&
                this.start.Value > this.end.Value)
            {
                throw new ServiceValidationException("Start date time cannot be bigger than end date time.");
            }
        }
    }
}