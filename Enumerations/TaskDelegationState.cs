// ---------------------------------------------------------------------------
// <copyright file="TaskDelegationState.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the TaskDelegationState enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    // This maps to the bogus TaskDelegationState in the EWS schema.
    // The schema enum has 6 values, but EWS should never return anything but
    // values between 0 and 3, so we should be safe without mappings for
    // EWS's Declined and Max values

    /// <summary>
    /// Defines the delegation state of a task.
    /// </summary>
    public enum TaskDelegationState
    {
        /// <summary>
        /// The task is not delegated
        /// </summary>
        NoDelegation, // Maps to NoMatch

        /// <summary>
        /// The task's delegation state is unknown.
        /// </summary>
        Unknown,      // Maps to OwnNew

        /// <summary>
        /// The task was delegated and the delegation was accepted.
        /// </summary>
        Accepted,     // Maps to Owned

        /// <summary>
        /// The task was delegated but the delegation was declined.
        /// </summary>
        Declined      // Maps to Accepted

        // The original Declined value has no mapping
        // The original Max value has no mapping
    }
}
