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