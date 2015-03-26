// ---------------------------------------------------------------------------
// <copyright file="GetServerTimeZonesResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetServerTimeZonesResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.ObjectModel;

    /// <summary>
    /// Represents the response to a GetServerTimeZones request.
    /// </summary>
    internal class GetServerTimeZonesResponse : ServiceResponse
    {
        private Collection<TimeZoneInfo> timeZones = new Collection<TimeZoneInfo>();

        /// <summary>
        /// Initializes a new instance of the <see cref="GetServerTimeZonesResponse"/> class.
        /// </summary>
        internal GetServerTimeZonesResponse()
            : base()
        {
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.TimeZoneDefinitions);

            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.TimeZoneDefinition))
                    {
                        TimeZoneDefinition timeZoneDefinition = new TimeZoneDefinition();
                        timeZoneDefinition.LoadFromXml(reader);

                        this.timeZones.Add(timeZoneDefinition.ToTimeZoneInfo(reader.Service));
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.TimeZoneDefinitions));
            }
        }

        /// <summary>
        /// Gets the time zones returned by the associated GetServerTimeZones request.
        /// </summary>
        /// <value>The time zones.</value>
        public Collection<TimeZoneInfo> TimeZones
        {
            get { return this.timeZones; }
        }
    }
}
