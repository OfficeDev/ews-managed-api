// ---------------------------------------------------------------------------
// <copyright file="AttendeeAvailability.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AttendeeAvailability class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Text;

    /// <summary>
    /// Represents the availability of an individual attendee.
    /// </summary>
    public sealed class AttendeeAvailability : ServiceResponse
    {
        private Collection<CalendarEvent> calendarEvents = new Collection<CalendarEvent>();
        private Collection<LegacyFreeBusyStatus> mergedFreeBusyStatus = new Collection<LegacyFreeBusyStatus>();
        private FreeBusyViewType viewType;
        private WorkingHours workingHours;

        /// <summary>
        /// Initializes a new instance of the <see cref="AttendeeAvailability"/> class.
        /// </summary>
        internal AttendeeAvailability()
            : base()
        {
        }

        /// <summary>
        /// Loads the free busy view from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="viewType">Type of free/busy view.</param>
        internal void LoadFreeBusyViewFromXml(EwsServiceXmlReader reader, FreeBusyViewType viewType)
        {
            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.FreeBusyView);

            string viewTypeString = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.FreeBusyViewType);

            this.viewType = (FreeBusyViewType)Enum.Parse(typeof(FreeBusyViewType), viewTypeString, false);

            do
            {
                reader.Read();

                if (reader.IsStartElement())
                {
                    switch (reader.LocalName)
                    {
                        case XmlElementNames.MergedFreeBusy:
                            string mergedFreeBusy = reader.ReadElementValue();

                            for (int i = 0; i < mergedFreeBusy.Length; i++)
                            {
                                this.mergedFreeBusyStatus.Add((LegacyFreeBusyStatus)Byte.Parse(mergedFreeBusy[i].ToString()));
                            }

                            break;
                        case XmlElementNames.CalendarEventArray:
                            do
                            {
                                reader.Read();

                                if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.CalendarEvent))
                                {
                                    CalendarEvent calendarEvent = new CalendarEvent();

                                    calendarEvent.LoadFromXml(reader, XmlElementNames.CalendarEvent);

                                    this.calendarEvents.Add(calendarEvent);
                                }
                            }
                            while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.CalendarEventArray));

                            break;
                        case XmlElementNames.WorkingHours:
                            this.workingHours = new WorkingHours();
                            this.workingHours.LoadFromXml(reader, reader.LocalName);

                            break;
                    }
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.FreeBusyView));
        }

        /// <summary>
        /// Gets a collection of calendar events for the attendee.
        /// </summary>
        public Collection<CalendarEvent> CalendarEvents
        {
            get { return this.calendarEvents; }
        }

        /// <summary>
        /// Gets the free/busy view type that wes retrieved for the attendee.
        /// </summary>
        public FreeBusyViewType ViewType
        {
            get { return this.viewType; }
        }

        /// <summary>
        /// Gets a collection of merged free/busy status for the attendee.
        /// </summary>
        public Collection<LegacyFreeBusyStatus> MergedFreeBusyStatus
        {
            get { return this.mergedFreeBusyStatus; }
        }

        /// <summary>
        /// Gets the working hours of the attendee.
        /// </summary>
        public WorkingHours WorkingHours
        {
            get { return this.workingHours; }
        }
    }
}