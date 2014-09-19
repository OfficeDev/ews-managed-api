// ---------------------------------------------------------------------------
// <copyright file="DateTimePrecision.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DateTimePrecision enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the precision for returned DateTime values
    /// </summary>
    public enum DateTimePrecision
    {
        /// <summary>
        /// Default value.  No SOAP header emitted.
        /// </summary>
        Default,

        /// <summary>
        /// Seconds
        /// </summary>
        Seconds,

        /// <summary>
        /// Milliseconds
        /// </summary>
        Milliseconds
    }
}
