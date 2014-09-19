// ---------------------------------------------------------------------------
// <copyright file="GetUserAvailabilityResults.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetUserAvailabilityResults class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Text;

    /// <summary>
    /// Represents the results of a GetUserAvailability operation.
    /// </summary>
    public sealed class GetUserAvailabilityResults
    {
        private ServiceResponseCollection<AttendeeAvailability> attendeesAvailability;
        private SuggestionsResponse suggestionsResponse;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetUserAvailabilityResults"/> class.
        /// </summary>
        internal GetUserAvailabilityResults()
        {
        }

        /// <summary>
        /// Gets or sets the suggestions response for the requested meeting time.
        /// </summary>
        internal SuggestionsResponse SuggestionsResponse
        {
            get { return this.suggestionsResponse; }
            set { this.suggestionsResponse = value; }
        }

        /// <summary>
        /// Gets a collection of AttendeeAvailability objects representing availability information for each of the specified attendees.
        /// </summary>
        public ServiceResponseCollection<AttendeeAvailability> AttendeesAvailability
        {
            get { return this.attendeesAvailability; }
            internal set { this.attendeesAvailability = value; }
        }

        /// <summary>
        /// Gets a collection of suggested meeting times for the specified time period.
        /// </summary>
        public Collection<Suggestion> Suggestions
        {
            get
            {
                if (this.suggestionsResponse == null)
                {
                    return null;
                }
                else
                {
                    this.suggestionsResponse.ThrowIfNecessary();

                    return this.suggestionsResponse.Suggestions;
                }
            }
        }
    }
}