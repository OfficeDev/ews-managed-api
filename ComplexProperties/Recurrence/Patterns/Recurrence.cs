// ---------------------------------------------------------------------------
// <copyright file="Recurrence.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the Recurrence class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.IO;
    using System.Text;

    /// <summary>
    /// Represents a recurrence pattern, as used by Appointment and Task items.
    /// </summary>
    public abstract partial class Recurrence : ComplexProperty
    {
        private DateTime? startDate;
        private int? numberOfOccurrences;
        private DateTime? endDate;

        /// <summary>
        /// Initializes a new instance of the <see cref="Recurrence"/> class.
        /// </summary>
        internal Recurrence()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Recurrence"/> class.
        /// </summary>
        /// <param name="startDate">The start date.</param>
        internal Recurrence(DateTime startDate)
            : this()
        {
            this.startDate = startDate;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <value>The name of the XML element.</value>
        internal abstract string XmlElementName { get; }

        /// <summary>
        /// Gets a value indicating whether this instance is regeneration pattern.
        /// </summary>
        /// <value>
        ///     <c>true</c> if this instance is regeneration pattern; otherwise, <c>false</c>.
        /// </value>
        internal virtual bool IsRegenerationPattern
        {
            get { return false; }
        }

        /// <summary>
        /// Write properties to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal virtual void InternalWritePropertiesToXml(EwsServiceXmlWriter writer)
        {
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override sealed void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Types, this.XmlElementName);
            this.InternalWritePropertiesToXml(writer);
            writer.WriteEndElement();

            RecurrenceRange range;

            if (!this.HasEnd)
            {
                range = new NoEndRecurrenceRange(this.StartDate);
            }
            else if (this.NumberOfOccurrences.HasValue)
            {
                range = new NumberedRecurrenceRange(this.StartDate, this.NumberOfOccurrences);
            }
            else
            {
                range = new EndDateRecurrenceRange(this.StartDate, this.EndDate.Value);
            }

            range.WriteToXml(writer, range.XmlElementName);
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

            jsonProperty.Add(JsonNames.RecurrencePattern, this.PatternToJson(service));
            jsonProperty.Add(JsonNames.RecurrenceRange, this.RangeToJson(service));

            return jsonProperty;
        }

        /// <summary>
        /// Ranges to json.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns></returns>
        private object RangeToJson(ExchangeService service)
        {
            RecurrenceRange range;

            if (!this.HasEnd)
            {
                range = new NoEndRecurrenceRange(this.StartDate);
            }
            else if (this.NumberOfOccurrences.HasValue)
            {
                range = new NumberedRecurrenceRange(this.StartDate, this.NumberOfOccurrences);
            }
            else
            {
                range = new EndDateRecurrenceRange(this.StartDate, this.EndDate.Value);
            }

            return range.InternalToJson(service);
        }

        /// <summary>
        /// Patterns to json.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns></returns>
        internal abstract JsonObject PatternToJson(ExchangeService service);

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            base.LoadFromJson(jsonProperty, service);
        }

        /// <summary>
        /// Gets a property value or throw if null.
        /// </summary>
        /// <typeparam name="T">Value type.</typeparam>
        /// <param name="value">The value.</param>
        /// <param name="name">The property name.</param>
        /// <returns>Property value</returns>
        internal T GetFieldValueOrThrowIfNull<T>(Nullable<T> value, string name) where T : struct
        {
            if (value.HasValue)
            {
                return value.Value;
            }
            else
            {
                throw new ServiceValidationException(
                                string.Format(Strings.PropertyValueMustBeSpecifiedForRecurrencePattern, name));
            }
        }

        /// <summary>
        /// Gets or sets the date and time when the recurrence start.
        /// </summary>
        public DateTime StartDate
        {
            get { return this.GetFieldValueOrThrowIfNull<DateTime>(this.startDate, "StartDate"); }
            set { this.startDate = value; }
        }

        /// <summary>
        /// Gets a value indicating whether the pattern has a fixed number of occurrences or an end date.
        /// </summary>
        public bool HasEnd
        {
            get { return this.numberOfOccurrences.HasValue || this.endDate.HasValue; }
        }

        /// <summary>
        /// Sets up this recurrence so that it never ends. Calling NeverEnds is equivalent to setting both NumberOfOccurrences and EndDate to null.
        /// </summary>
        public void NeverEnds()
        {
            this.numberOfOccurrences = null;
            this.endDate = null;
            this.Changed();
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void InternalValidate()
        {
            base.InternalValidate();

            if (!this.startDate.HasValue)
            {
                throw new ServiceValidationException(Strings.RecurrencePatternMustHaveStartDate);
            }
        }

        /// <summary>
        /// Gets or sets the number of occurrences after which the recurrence ends. Setting NumberOfOccurrences resets EndDate.
        /// </summary>
        public int? NumberOfOccurrences
        {
            get
            {
                return this.numberOfOccurrences;
            }

            set
            {
                if (value < 1)
                {
                    throw new ArgumentException(Strings.NumberOfOccurrencesMustBeGreaterThanZero);
                }

                this.SetFieldValue<int?>(ref this.numberOfOccurrences, value);
                this.endDate = null;
            }
        }

        /// <summary>
        /// Gets or sets the date after which the recurrence ends. Setting EndDate resets NumberOfOccurrences.
        /// </summary>
        public DateTime? EndDate
        {
            get
            {
                return this.endDate;
            }

            set
            {
                this.SetFieldValue<DateTime?>(ref this.endDate, value);
                this.numberOfOccurrences = null;
            }
        }
        
        /// <summary>
        /// Compares two objects by converting them to JSON and comparing their string values 
        /// </summary>
        /// <param name="otherRecurrence">object to compare to</param>
        /// <returns>true if the objects serialize to the same string</returns>
        public bool IsSame(Recurrence otherRecurrence)
        {
            if (otherRecurrence == null)
            {
                return false;
            }

            string jsonString;
            using (MemoryStream memoryStream = new MemoryStream())
            {
                ((JsonObject)this.InternalToJson(null)).SerializeToJson(memoryStream);
                memoryStream.Position = 0;
                using (StreamReader reader = new StreamReader(memoryStream))
                {
                    jsonString = reader.ReadToEnd();
                }
            }

            string otherJsonString;
            using (MemoryStream memoryStream = new MemoryStream())
            {
                ((JsonObject)otherRecurrence.InternalToJson(null)).SerializeToJson(memoryStream);
                memoryStream.Position = 0;
                using (StreamReader reader = new StreamReader(memoryStream))
                {
                    otherJsonString = reader.ReadToEnd();
                }
            }

            return String.Equals(jsonString, otherJsonString, StringComparison.Ordinal);
        }
    }
}
