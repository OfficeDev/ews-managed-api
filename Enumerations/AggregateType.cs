// ---------------------------------------------------------------------------
// <copyright file="AggregateType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AggregateType enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the type of aggregation to perform.
    /// </summary>
    public enum AggregateType
    {
        /// <summary>
        /// The maximum value is calculated.
        /// </summary>
        Minimum,

        /// <summary>
        /// The minimum value is calculated.
        /// </summary>
        Maximum
    }
}
