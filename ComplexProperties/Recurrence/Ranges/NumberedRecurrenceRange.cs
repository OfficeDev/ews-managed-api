// ---------------------------------------------------------------------------
// <copyright file="NumberedRecurrenceRange.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the NumberedRecurrenceRange class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    internal sealed class NumberedRecurrenceRange : RecurrenceRange
    {
        private int? numberOfOccurrences;

        /// <summary>
        /// Initializes a new instance of the <see cref="NumberedRecurrenceRange"/> class.
        /// </summary>
        public NumberedRecurrenceRange()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="NumberedRecurrenceRange"/> class.
        /// </summary>
        /// <param name="startDate">The start date.</param>
        /// <param name="numberOfOccurrences">The number of occurrences.</param>
        public NumberedRecurrenceRange(DateTime startDate, int? numberOfOccurrences)
            : base(startDate)
        {
            this.numberOfOccurrences = numberOfOccurrences;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <value>The name of the XML element.</value>
        internal override string XmlElementName
        {
            get { return XmlElementNames.NumberedRecurrence; }
        }

        /// <summary>
        /// Setups the recurrence.
        /// </summary>
        /// <param name="recurrence">The recurrence.</param>
        internal override void SetupRecurrence(Recurrence recurrence)
        {
            base.SetupRecurrence(recurrence);

            recurrence.NumberOfOccurrences = this.NumberOfOccurrences;
        }

        /// <summary>
        /// Writes the elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            base.WriteElementsToXml(writer);

            if (this.NumberOfOccurrences.HasValue)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.NumberOfOccurrences,
                    this.NumberOfOccurrences);
            }
        }

        /// <summary>
        /// Adds the properties to json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void AddPropertiesToJson(JsonObject jsonProperty, ExchangeService service)
        {
            base.AddPropertiesToJson(jsonProperty, service);

            jsonProperty.Add(XmlElementNames.NumberOfOccurrences, this.NumberOfOccurrences);
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
                    case XmlElementNames.NumberOfOccurrences:
                        this.numberOfOccurrences = reader.ReadElementValue<int>();
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
                    case XmlElementNames.NumberOfOccurrences:
                        this.numberOfOccurrences = jsonProperty.ReadAsInt(key);
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Gets or sets the number of occurrences.
        /// </summary>
        /// <value>The number of occurrences.</value>
        public int? NumberOfOccurrences
        {
            get
            {
                return this.numberOfOccurrences;
            }

            set
            {
                this.SetFieldValue<int?>(ref this.numberOfOccurrences, value);
            }
        }
    }
}
