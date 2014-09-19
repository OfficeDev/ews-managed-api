// ---------------------------------------------------------------------------
// <copyright file="AvailabilityData.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AvailabilityData enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the type of data that can be requested via GetUserAvailability.
    /// </summary>
    public enum AvailabilityData
    {
        /// <summary>
        /// Only return free/busy data.
        /// </summary>
        FreeBusy,

        /// <summary>
        /// Only return suggestions.
        /// </summary>
        Suggestions,

        /// <summary>
        /// Return both free/busy data and suggestions.
        /// </summary>
        FreeBusyAndSuggestions
    }
}
