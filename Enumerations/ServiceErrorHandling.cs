// ---------------------------------------------------------------------------
// <copyright file="ServiceErrorHandling.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ServiceErrorHandling enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the type of error handling used for service method calls. 
    /// </summary>
    internal enum ServiceErrorHandling
    {
        /// <summary>
        /// Service method should return the error(s).
        /// </summary>
        ReturnErrors,

        /// <summary>
        /// Service method should throw exception when error occurs.
        /// </summary>
        ThrowOnError
    }
}
