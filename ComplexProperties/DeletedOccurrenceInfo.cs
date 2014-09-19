// ---------------------------------------------------------------------------
// <copyright file="DeletedOccurrenceInfo.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DeletedOccurrenceInfo class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Encapsulates information on the deleted occurrence of a recurring appointment.
    /// </summary>
    public class DeletedOccurrenceInfo : ComplexProperty
    {
        /// <summary>
        /// The original start date and time of the deleted occurrence.
        /// </summary>
        /// <remarks>
        /// The EWS schema contains a Start property for deleted occurrences but it's
        /// really the original start date and time of the occurrence.
        /// </remarks>
        private DateTime originalStart;

        /// <summary>
        /// Initializes a new instance of the <see cref="DeletedOccurrenceInfo"/> class.
        /// </summary>
        internal DeletedOccurrenceInfo()
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
                case XmlElementNames.Start:
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
        /// <param name="service"></param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            if (jsonProperty.ContainsKey(XmlElementNames.Start))
            {
                this.originalStart = service.ConvertUniversalDateTimeStringToLocalDateTime(jsonProperty.ReadAsString(XmlElementNames.Start)).Value;
            }
        }

        /// <summary>
        /// Gets the original start date and time of the deleted occurrence.
        /// </summary>
        public DateTime OriginalStart
        {
            get { return this.originalStart; }
        }
    }
}
