// ---------------------------------------------------------------------------
// <copyright file="OccurrenceInfo.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the OccurrenceInfo class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Encapsulates information on the occurrence of a recurring appointment.
    /// </summary>
    public sealed class OccurrenceInfo : ComplexProperty
    {
        private ItemId itemId;
        private DateTime start;
        private DateTime end;
        private DateTime originalStart;

        /// <summary>
        /// Initializes a new instance of the <see cref="OccurrenceInfo"/> class.
        /// </summary>
        internal OccurrenceInfo()
        {
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
                case XmlElementNames.ItemId:
                    this.itemId = new ItemId();
                    this.itemId.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.Start:
                    this.start = reader.ReadElementValueAsDateTime().Value;
                    return true;
                case XmlElementNames.End:
                    this.end = reader.ReadElementValueAsDateTime().Value;
                    return true;
                case XmlElementNames.OriginalStart:
                    this.originalStart = reader.ReadElementValueAsDateTime().Value;
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
                    case XmlElementNames.ItemId:
                        this.itemId = new ItemId();
                        this.itemId.LoadFromJson(jsonProperty.ReadAsJsonObject(key), service);
                        break;
                    case XmlElementNames.Start:
                        this.start = service.ConvertUniversalDateTimeStringToLocalDateTime(jsonProperty.ReadAsString(key)).Value;
                        break;
                    case XmlElementNames.End:
                        this.end = service.ConvertUniversalDateTimeStringToLocalDateTime(jsonProperty.ReadAsString(key)).Value;
                        break;
                    case XmlElementNames.OriginalStart:
                        this.originalStart = service.ConvertUniversalDateTimeStringToLocalDateTime(jsonProperty.ReadAsString(key)).Value;
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Gets the Id of the occurrence.
        /// </summary>
        public ItemId ItemId
        {
            get { return this.itemId; }
        }

        /// <summary>
        /// Gets the start date and time of the occurrence.
        /// </summary>
        public DateTime Start
        {
            get { return this.start; }
        }

        /// <summary>
        /// Gets the end date and time of the occurrence.
        /// </summary>
        public DateTime End
        {
            get { return this.end; }
        }

        /// <summary>
        /// Gets the original start date and time of the occurrence.
        /// </summary>
        public DateTime OriginalStart
        {
            get { return this.originalStart; }
        }
    }
}
