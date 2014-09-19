// ---------------------------------------------------------------------------
// <copyright file="LegacyFreeBusyStatus.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the LegacyFreeBusyStatus enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the legacy free/busy status associated with an appointment.
    /// </summary>
    public enum LegacyFreeBusyStatus
    {
        /// <summary>
        /// The time slot associated with the appointment appears as free.
        /// </summary>
        Free = 0,

        /// <summary>
        /// The time slot associated with the appointment appears as tentative.
        /// </summary>
        Tentative = 1,

        /// <summary>
        /// The time slot associated with the appointment appears as busy.
        /// </summary>
        Busy = 2,

        /// <summary>
        /// The time slot associated with the appointment appears as Out of Office.
        /// </summary>
        OOF = 3,

        /// <summary>
        /// The time slot associated with the appointment appears as working else where.
        /// </summary>
        WorkingElsewhere = 4,

        /// <summary>
        /// No free/busy status is associated with the appointment.
        /// </summary>
        NoData = 5
    }
}
