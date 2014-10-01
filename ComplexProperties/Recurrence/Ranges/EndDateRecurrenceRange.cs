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
// <summary>Defines the EndDateRecurrenceRange class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents recurrent range with an end date.
    /// </summary>
    internal sealed class EndDateRecurrenceRange : RecurrenceRange
    {
        private DateTime endDate;

        /// <summary>
        /// Initializes a new instance of the <see cref="EndDateRecurrenceRange"/> class.
        /// </summary>
        public EndDateRecurrenceRange()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EndDateRecurrenceRange"/> class.
        /// </summary>
        /// <param name="startDate">The start date.</param>
        /// <param name="endDate">The end date.</param>
        public EndDateRecurrenceRange(DateTime startDate, DateTime endDate)
            : base(startDate)
        {
            this.endDate = endDate;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <value>The name of the XML element.</value>
        internal override string XmlElementName
        {
            get { return XmlElementNames.EndDateRecurrence; }
        }

        /// <summary>
        /// Setups the recurrence.
        /// </summary>
        /// <param name="recurrence">The recurrence.</param>
        internal override void SetupRecurrence(Recurrence recurrence)
        {
            base.SetupRecurrence(recurrence);

            recurrence.EndDate = this.EndDate;
        }

        /// <summary>
        /// Writes the elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            base.WriteElementsToXml(writer);

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.EndDate,
                EwsUtilities.DateTimeToXSDate(this.EndDate));
        }

        /// <summary>
        /// Adds the properties to json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void AddPropertiesToJson(JsonObject jsonProperty, ExchangeService service)
        {
            base.AddPropertiesToJson(jsonProperty, service);

            jsonProperty.Add(XmlElementNames.EndDate, EwsUtilities.DateTimeToXSDate(this.EndDate));
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            if (base.TryReadElementFromXml(reader))
            {
                return true;
            }
            else
            {
                switch (reader.LocalName)
                {
                    case XmlElementNames.EndDate:
                        this.endDate = reader.ReadElementValueAsDateTime().Value;
                        return true;
                    default:
                        return false;
                }
            }
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            base.LoadFromJson(jsonProperty, service);

            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlElementNames.EndDate:
                        this.endDate = service.ConvertStartDateToUnspecifiedDateTime(jsonProperty.ReadAsString(key)).Value;
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Gets or sets the end date.
        /// </summary>
        /// <value>The end date.</value>
        public DateTime EndDate
        {
            get { return this.endDate; }
            set { this.SetFieldValue<DateTime>(ref this.endDate, value); }
        }
    }
}
