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