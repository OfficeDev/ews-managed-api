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
        /// Checks if two recurrence objects are identical. 
        /// </summary>
        /// <param name="otherRecurrence">The recurrence to compare this one to.</param>
        /// <returns>true if the two recurrences are identical, false otherwise.</returns>
        public virtual bool IsSame(Recurrence otherRecurrence)
        {
            if (otherRecurrence == null)
            {
                return false;
            }

            return (this.GetType() == otherRecurrence.GetType() &&
                    this.numberOfOccurrences == otherRecurrence.numberOfOccurrences &&
                    this.endDate == otherRecurrence.endDate &&
                    this.startDate == otherRecurrence.startDate);
        }
    }
}