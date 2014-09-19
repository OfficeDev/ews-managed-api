// ---------------------------------------------------------------------------
// <copyright file="MeetingRequestsDeliveryScope.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MeetingRequestsDeliveryScope enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines how meeting requests are sent to delegates.
    /// </summary>
    public enum MeetingRequestsDeliveryScope
    {
        /// <summary>
        /// Meeting requests are sent to delegates only.
        /// </summary>
        DelegatesOnly,

        /// <summary>
        /// Meeting requests are sent to delegates and to the owner of the mailbox.
        /// </summary>
        DelegatesAndMe,

        /// <summary>
        /// Meeting requests are sent to delegates and informational messages are sent to the owner of the mailbox.
        /// </summary>
        DelegatesAndSendInformationToMe,

        /// <summary>
        /// Meeting requests are not sent to delegates.  This value is supported only for Exchange 2010 SP1 or later
        /// server versions.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2010_SP1)]
        NoForward
    }
}
