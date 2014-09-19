// ---------------------------------------------------------------------------
// <copyright file="DelegateInformation.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DelegateInformation class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.Generic;
    using System.Collections.ObjectModel;

    /// <summary>
    /// Represents the results of a GetDelegates operation.
    /// </summary>
    public sealed class DelegateInformation
    {
        #region Private members

        private Collection<DelegateUserResponse> delegateUserResponses;
        private MeetingRequestsDeliveryScope meetingReqestsDeliveryScope;

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a DelegateInformation object
        /// </summary>
        /// <param name="delegateUserResponses">List of DelegateUserResponses from a GetDelegates request</param>
        /// <param name="meetingReqestsDeliveryScope">MeetingRequestsDeliveryScope from a GetDelegates request.</param>
        internal DelegateInformation(IList<DelegateUserResponse> delegateUserResponses, MeetingRequestsDeliveryScope meetingReqestsDeliveryScope)
        {
            this.delegateUserResponses = new Collection<DelegateUserResponse>(delegateUserResponses);
            this.meetingReqestsDeliveryScope = meetingReqestsDeliveryScope;
        }

        #endregion

        #region Public Properties

        /// <summary>
        /// Gets a list of responses for each of the delegate users concerned by the operation.
        /// </summary>
        public Collection<DelegateUserResponse> DelegateUserResponses
        {
            get { return this.delegateUserResponses; }
        }

        /// <summary>
        /// Gets a value indicating if and how meeting requests are delivered to delegates.
        /// </summary>
        public MeetingRequestsDeliveryScope MeetingRequestsDeliveryScope
        {
            get { return this.meetingReqestsDeliveryScope; }
        }

        #endregion
    }
}
