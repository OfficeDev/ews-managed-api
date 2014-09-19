// ---------------------------------------------------------------------------
// <copyright file="Importance.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the Importance enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the importance of an item.
    /// </summary>
    public enum Importance
    {
        /// <summary>
        /// Low importance.
        /// </summary>
        Low,

        /// <summary>
        /// Normal importance.
        /// </summary>
        Normal,

        /// <summary>
        /// High importance.
        /// </summary>
        High
    }
}
