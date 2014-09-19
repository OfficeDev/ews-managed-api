// ---------------------------------------------------------------------------
// <copyright file="AffectedTaskOccurrence.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//------------------------------------------------------------------------------
// <summary>Defines the AffectedTaskOccurrence enumeration.</summary>
//------------------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Indicates which occurrence of a recurring task should be deleted.
    /// </summary>
    public enum AffectedTaskOccurrence
    {
        /// <summary>
        /// All occurrences of the recurring task will be deleted.
        /// </summary>
        AllOccurrences,

        /// <summary>
        /// Only the current occurrence of the recurring task will be deleted.
        /// </summary>
        SpecifiedOccurrenceOnly
    }
}