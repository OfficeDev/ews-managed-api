#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

//-----------------------------------------------------------------------
// <summary>Defines the AvailabilityOptions class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the options of a GetAvailability request.
    /// </summary>
    public sealed class AvailabilityOptions
    {
        private int mergedFreeBusyInterval = 30;
        private FreeBusyViewType requestedFreeBusyView = FreeBusyViewType.Detailed;
        private int goodSuggestionThreshold = 25;
        private int maximumSuggestionsPerDay = 10;
        private int maximumNonWorkHoursSuggestionsPerDay = 0;
        private int meetingDuration = 60;
        private SuggestionQuality minimumSuggestionQuality = SuggestionQuality.Fair;
        private TimeWindow detailedSuggestionsWindow;
        private DateTime? currentMeetingTime;
        private string globalObjectId;

        /// <summary>
        /// Validates this instance against the specified time window.
        /// </summary>
        /// <param name="timeWindow">The time window.</param>
        internal void Validate(TimeSpan timeWindow)
        {
            if (TimeSpan.FromMinutes(this.MergedFreeBusyInterval) > timeWindow)
            {
                throw new ArgumentException(Strings.MergedFreeBusyIntervalMustBeSmallerThanTimeWindow, "MergedFreeBusyInterval");
            }

            EwsUtilities.ValidateParamAllowNull(this.DetailedSuggestionsWindow, "DetailedSuggestionsWindow");
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="request">The request being emitted.</param>
        internal void WriteToXml(EwsServiceXmlWriter writer, GetUserAvailabilityRequest request)
        {
            if (request.IsFreeBusyViewRequested)
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.FreeBusyViewOptions);

                request.TimeWindow.WriteToXmlUnscopedDatesOnly(writer, XmlElementNames.TimeWindow);

                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.MergedFreeBusyIntervalInMinutes,
                    this.MergedFreeBusyInterval);

                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.RequestedView,
                    this.RequestedFreeBusyView);

                writer.WriteEndElement(); // FreeBusyViewOptions
            }

            if (request.IsSuggestionsViewRequested)
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.SuggestionsViewOptions);

                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.GoodThreshold,
                    this.GoodSuggestionThreshold);

                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.MaximumResultsByDay,
                    this.MaximumSuggestionsPerDay);

                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.MaximumNonWorkHourResultsByDay,
                    this.MaximumNonWorkHoursSuggestionsPerDay);

                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.MeetingDurationInMinutes,
                    this.MeetingDuration);

                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.MinimumSuggestionQuality,
                    this.MinimumSuggestionQuality);

                TimeWindow timeWindowToSerialize = this.DetailedSuggestionsWindow == null ?
                    request.TimeWindow :
                    this.DetailedSuggestionsWindow;

                timeWindowToSerialize.WriteToXmlUnscopedDatesOnly(writer, XmlElementNames.DetailedSuggestionsWindow);

                if (this.CurrentMeetingTime.HasValue)
                {
                    writer.WriteElementValue(
                        XmlNamespace.Types,
                        XmlElementNames.CurrentMeetingTime,
                        this.CurrentMeetingTime.Value);
                }

                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.GlobalObjectId,
                    this.GlobalObjectId);

                writer.WriteEndElement(); // SuggestionsViewOptions
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AvailabilityOptions"/> class.
        /// </summary>
        public AvailabilityOptions()
        {
        }

        /// <summary>
        /// Gets or sets the time difference between two successive slots in a FreeBusyMerged view.
        /// MergedFreeBusyInterval must be between 5 and 1440. The default value is 30.
        /// </summary>
        public int MergedFreeBusyInterval
        {
            get
            {
                return this.mergedFreeBusyInterval;
            }

            set
            {
                if (value < 5 || value > 1440)
                {
                    throw new ArgumentException(
                        string.Format(
                            Strings.InvalidPropertyValueNotInRange,
                            "MergedFreeBusyInterval",
                            5,
                            1440));
                }

                this.mergedFreeBusyInterval = value;
            }
        }

        /// <summary>
        /// Gets or sets the requested type of free/busy view. The default value is FreeBusyViewType.Detailed.
        /// </summary>
        public FreeBusyViewType RequestedFreeBusyView
        {
            get { return this.requestedFreeBusyView; }
            set { this.requestedFreeBusyView = value; }
        }

        /// <summary>
        /// Gets or sets the percentage of attendees that must have the time period open for the time period to qualify as a good suggested meeting time.
        /// GoodSuggestionThreshold must be between 1 and 49. The default value is 25.
        /// </summary>
        public int GoodSuggestionThreshold
        {
            get
            {
                return this.goodSuggestionThreshold;
            }

            set
            {
                if (value < 1 || value > 49)
                {
                    throw new ArgumentException(
                        string.Format(
                            Strings.InvalidPropertyValueNotInRange,
                            "GoodSuggestionThreshold",
                            1,
                            49));
                }

                this.goodSuggestionThreshold = value;
            }
        }

        /// <summary>
        /// Gets or sets the number of suggested meeting times that should be returned per day.
        /// MaximumSuggestionsPerDay must be between 0 and 48. The default value is 10.
        /// </summary>
        public int MaximumSuggestionsPerDay
        {
            get
            {
                return this.maximumSuggestionsPerDay;
            }

            set
            {
                if (value < 0 || value > 48)
                {
                    throw new ArgumentException(
                        string.Format(
                            Strings.InvalidPropertyValueNotInRange,
                            "MaximumSuggestionsPerDay",
                            0,
                            48));
                }

                this.maximumSuggestionsPerDay = value;
            }
        }

        /// <summary>
        /// Gets or sets the number of suggested meeting times outside regular working hours per day.
        /// MaximumNonWorkHoursSuggestionsPerDay must be between 0 and 48. The default value is 0.
        /// </summary>
        public int MaximumNonWorkHoursSuggestionsPerDay
        {
            get
            {
                return this.maximumNonWorkHoursSuggestionsPerDay;
            }

            set
            {
                if (value < 0 || value > 48)
                {
                    throw new ArgumentException(
                        string.Format(
                            Strings.InvalidPropertyValueNotInRange,
                            "MaximumNonWorkHoursSuggestionsPerDay",
                            0,
                            48));
                }

                this.maximumNonWorkHoursSuggestionsPerDay = value;
            }
        }

        /// <summary>
        /// Gets or sets the duration, in minutes, of the meeting for which to obtain suggestions.
        /// MeetingDuration must be between 30 and 1440. The default value is 60.
        /// </summary>
        public int MeetingDuration
        {
            get
            {
                return this.meetingDuration;
            }

            set
            {
                if (value < 30 || value > 1440)
                {
                    throw new ArgumentException(
                        string.Format(
                            Strings.InvalidPropertyValueNotInRange,
                            "MeetingDuration",
                            30,
                            1440));
                }

                this.meetingDuration = value;
            }
        }

        /// <summary>
        /// Gets or sets the minimum quality of suggestions that should be returned.
        /// The default is SuggestionQuality.Fair.
        /// </summary>
        public SuggestionQuality MinimumSuggestionQuality
        {
            get { return this.minimumSuggestionQuality; }
            set { this.minimumSuggestionQuality = value; }
        }

        /// <summary>
        /// Gets or sets the time window for which detailed information about suggested meeting times should be returned.
        /// </summary>
        public TimeWindow DetailedSuggestionsWindow
        {
            get { return this.detailedSuggestionsWindow; }
            set { this.detailedSuggestionsWindow = value; }
        }

        /// <summary>
        /// Gets or sets the start time of a meeting that you want to update with the suggested meeting times.
        /// </summary>
        public DateTime? CurrentMeetingTime
        {
            get { return this.currentMeetingTime; }
            set { this.currentMeetingTime = value; }
        }

        /// <summary>
        /// Gets or sets the global object Id of a meeting that will be modified based on the data returned by GetUserAvailability.
        /// </summary>
        public string GlobalObjectId
        {
            get { return this.globalObjectId; }
            set { this.globalObjectId = value; }
        }
    }
}
