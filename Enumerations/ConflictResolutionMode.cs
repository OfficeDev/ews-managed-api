// ---------------------------------------------------------------------------
// <copyright file="ConflictResolutionMode.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ConflictResolution enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines how conflict resolutions are handled in update operations.
    /// </summary>
    public enum ConflictResolutionMode
    {
        /// <summary>
        /// Local property changes are discarded.
        /// </summary>
        NeverOverwrite,

        /// <summary>
        /// Local property changes are applied to the server unless the server-side copy is more recent than the local copy.
        /// </summary>
        AutoResolve,

        /// <summary>
        /// Local property changes overwrite server-side changes. 
        /// </summary>
        AlwaysOverwrite
    }
}
