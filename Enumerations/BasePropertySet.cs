// ---------------------------------------------------------------------------
// <copyright file="BasePropertySet.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the BasePropertySet enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines base property sets that are used as the base for custom property sets.
    /// </summary>
    public enum BasePropertySet
    {
        /// <summary>
        /// Only includes the Id of items and folders.
        /// </summary>
        IdOnly,

        /// <summary>
        /// Includes all the first class properties of items and folders.
        /// </summary>
        FirstClassProperties
    }
}