// ---------------------------------------------------------------------------
// <copyright file="TimeWindow.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the TimeWindow class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Text;

    /// <summary>
    /// Represents a time period.
    /// </summary>
    public sealed class TimeWindow : ISelfValidate
    {
        private DateTime startTime;
        private DateTime endTime;

        /// <summary>
        /// Initializes a new instance of the <see cref="TimeWindow"/> class.
        /// </summary>
        internal TimeWindow()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TimeWindow"/> class.
        /// </summary>
        /// <param name="startTime">The start date and time.</param>
        /// <param name="endTime">The end date and time.</param>
        public TimeWindow(DateTime startTime, DateTime endTime)
            : this()
        {
            this.startTime = startTime;
            this.endTime = endTime;
        }

        /// <summary>
        /// Gets or sets the start date and time.
        /// </summary>
        public DateTime StartTime
        {
            get { return this.startTime; }
            set { this.startTime = value; }
        }

        /// <summary>
        /// Gets or sets the end date and time.
        /// </summary>
        public DateTime EndTime
        {
            get { return this.endTime; }
            set { this.endTime = value; }
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadFromXml(EwsServiceXmlReader reader)
        {
            reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, XmlElementNames.Duration);

            this.startTime = reader.ReadElementValueAsDateTime(XmlNamespace.Types, XmlElementNames.StartTime).Value;
            this.endTime = reader.ReadElementValueAsDateTime(XmlNamespace.Types, XmlElementNames.EndTime).Value;

            reader.ReadEndElement(XmlNamespace.Types, XmlElementNames.Duration);
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        /// <param name="service">The service.</param>
        internal void LoadFromJson(JsonObject jsonObject, ExchangeService service)
        {
            this.startTime = service.ConvertUniversalDateTimeStringToLocalDateTime(jsonObject.ReadAsString(XmlElementNames.StartTime)).Value;
            this.endTime = service.ConvertUniversalDateTimeStringToLocalDateTime(jsonObject.ReadAsString(XmlElementNames.EndTime)).Value;
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="startTime">The start time.</param>
        /// <param name="endTime">The end time.</param>
        private static void WriteToXml(
            EwsServiceXmlWriter writer,
            string xmlElementName,
            object startTime,
            object endTime)
        {
            writer.WriteStartElement(XmlNamespace.Types, xmlElementName);

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.StartTime,
                startTime);

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.EndTime,
                endTime);

            writer.WriteEndElement(); // xmlElementName
        }

        /// <summary>
        /// Writes to XML without scoping the dates and without emitting times.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal void WriteToXmlUnscopedDatesOnly(EwsServiceXmlWriter writer, string xmlElementName)
        {
            const string DateOnlyFormat = "yyyy-MM-ddT00:00:00";

            TimeWindow.WriteToXml(
                writer,
                xmlElementName,
                this.StartTime.ToString(DateOnlyFormat, CultureInfo.InvariantCulture),
                this.EndTime.ToString(DateOnlyFormat, CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal void WriteToXml(EwsServiceXmlWriter writer, string xmlElementName)
        {
            TimeWindow.WriteToXml(
                writer,
                xmlElementName,
                this.StartTime,
                this.EndTime);
        }

        /// <summary>
        /// Serializes to json.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns></returns>
        internal JsonObject InternalToJson(ExchangeService service)
        {
            JsonObject jsonProperty = new JsonObject();

            jsonProperty.Add(XmlElementNames.StartTime, EwsUtilities.DateTimeToXSDateTime(startTime));
            jsonProperty.Add(XmlElementNames.EndTime, EwsUtilities.DateTimeToXSDateTime(endTime));

            return jsonProperty;
        }

        /// <summary>
        /// Gets the duration.
        /// </summary>
        internal TimeSpan Duration
        {
            get { return this.endTime - this.startTime; }
        }

        #region ISelfValidate Members

        /// <summary>
        /// Validates this instance.
        /// </summary>
        void ISelfValidate.Validate()
        {
            if (this.StartTime >= this.EndTime)
            {
                throw new ArgumentException(Strings.TimeWindowStartTimeMustBeGreaterThanEndTime);
            }
        }

        #endregion
    }
}
