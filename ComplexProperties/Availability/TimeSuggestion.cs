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
    /// Represents an availability time suggestion.
    /// </summary>
    public sealed class TimeSuggestion : ComplexProperty
    {
        private DateTime meetingTime;
        private bool isWorkTime;
        private SuggestionQuality quality;
        private Collection<Conflict> conflicts = new Collection<Conflict>();

        /// <summary>
        /// Initializes a new instance of the <see cref="TimeSuggestion"/> class.
        /// </summary>
        internal TimeSuggestion()
            : base()
        {
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if appropriate element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.MeetingTime:
                    this.meetingTime = reader.ReadElementValueAsUnbiasedDateTimeScopedToServiceTimeZone();
                    return true;
                case XmlElementNames.IsWorkTime:
                    this.isWorkTime = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.SuggestionQuality:
                    this.quality = reader.ReadElementValue<SuggestionQuality>();
                    return true;
                case XmlElementNames.AttendeeConflictDataArray:
                    if (!reader.IsEmptyElement)
                    {
                        do
                        {
                            reader.Read();

                            if (reader.IsStartElement())
                            {
                                Conflict conflict = null;

                                switch (reader.LocalName)
                                {
                                    case XmlElementNames.UnknownAttendeeConflictData:
                                        conflict = new Conflict(ConflictType.UnknownAttendeeConflict);
                                        break;
                                    case XmlElementNames.TooBigGroupAttendeeConflictData:
                                        conflict = new Conflict(ConflictType.GroupTooBigConflict);
                                        break;
                                    case XmlElementNames.IndividualAttendeeConflictData:
                                        conflict = new Conflict(ConflictType.IndividualAttendeeConflict);
                                        break;
                                    case XmlElementNames.GroupAttendeeConflictData:
                                        conflict = new Conflict(ConflictType.GroupConflict);
                                        break;
                                    default:
                                        EwsUtilities.Assert(
                                            false,
                                            "TimeSuggestion.TryReadElementFromXml",
                                            string.Format("The {0} element name does not map to any AttendeeConflict descendant.", reader.LocalName));

                                        // The following line to please the compiler
                                        break;
                                }

                                conflict.LoadFromXml(reader, reader.LocalName);

                                this.conflicts.Add(conflict);
                            }
                        }
                        while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.AttendeeConflictDataArray));
                    }

                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Gets the suggested time.
        /// </summary>
        public DateTime MeetingTime
        {
            get { return this.meetingTime; }
        }

        /// <summary>
        /// Gets a value indicating whether the suggested time is within working hours.
        /// </summary>
        public bool IsWorkTime
        {
            get { return this.isWorkTime; }
        }

        /// <summary>
        /// Gets the quality of the suggestion.
        /// </summary>
        public SuggestionQuality Quality
        {
            get { return this.quality; }
        }

        /// <summary>
        /// Gets a collection of conflicts at the suggested time.
        /// </summary>
        public Collection<Conflict> Conflicts
        {
            get { return this.conflicts; }
        }
    }
}