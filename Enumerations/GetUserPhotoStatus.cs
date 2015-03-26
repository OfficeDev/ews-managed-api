// ---------------------------------------------------------------------------
// <copyright file="GetUserPhotoStatus.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetUserPhotoStatus enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the response types from a GetUserPhoto request
    /// </summary>
    public enum GetUserPhotoStatus
    {
        /// <summary>
        /// The photo was successfully returned
        /// </summary>
        PhotoReturned,

        /// <summary>
        /// The photo has not changed since it was last obtained
        /// </summary>
        PhotoUnchanged,

        /// <summary>
        /// The photo or user was not found on the server
        /// </summary>
        PhotoOrUserNotFound,
    }
}
