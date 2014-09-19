// ---------------------------------------------------------------------------
// <copyright file="RulePredicateDateRange.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the RulePredicateDateRange class.</summary>
//-----------------------------------------------------------------------

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
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlElementNames.StartDateTime:
                        this.start = service.ConvertUniversalDateTimeStringToLocalDateTime(jsonProperty.ReadAsString(key));
                        break;
                    case XmlElementNames.EndDateTime:
                        this.end = service.ConvertUniversalDateTimeStringToLocalDateTime(jsonProperty.ReadAsString(key));
                        break;
                    default:
                        break;
                }
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
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object InternalToJson(ExchangeService service)
        {
            JsonObject jsonProperty = new JsonObject();

            if (this.Start.HasValue)
            {
                jsonProperty.Add(XmlElementNames.StartDateTime, service.ConvertDateTimeToUniversalDateTimeString(this.Start.Value));
            }
            if (this.End.HasValue)
            {
                jsonProperty.Add(XmlElementNames.EndDateTime, service.ConvertDateTimeToUniversalDateTimeString(this.End.Value));
            }

            return jsonProperty;
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
