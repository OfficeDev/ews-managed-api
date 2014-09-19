// ---------------------------------------------------------------------------
// <copyright file="OffsetBasePoint.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the OffsetBasePoint enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the offset's base point in a paged view.
    /// </summary>
    public enum OffsetBasePoint
    {
        /// <summary>
        /// The offset is from the beginning of the view.
        /// </summary>
        Beginning,

        /// <summary>
        /// The offset is from the end of the view.
        /// </summary>
        End
    }
}
