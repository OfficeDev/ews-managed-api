// ---------------------------------------------------------------------------
// <copyright file="StandardUser.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the StandardUser enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines a standard delegate user.
    /// </summary>
    public enum StandardUser
    {
        /// <summary>
        /// The Default delegate user, used to define default delegation permissions.
        /// </summary>
        Default,

        /// <summary>
        /// The Anonymous delegate user, used to define delegate permissions for unauthenticated users.
        /// </summary>
        Anonymous
    }
}
