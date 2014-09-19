// ---------------------------------------------------------------------------
// <copyright file="LocationSource.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    // System Dependencies
    using System.Runtime.Serialization;

    /// <summary>
    /// Source of resolution.
    /// </summary>
    public enum LocationSource
    {
        /// <summary>Unresolved</summary>
        None = 0,

        /// <summary>Resolved by external location services (such as Bing, Google, etc)</summary>
        LocationServices = 1,

        /// <summary>Resolved by external phonebook services (such as Bing, Google, etc)</summary>
        PhonebookServices = 2,

        /// <summary>Revolved by a GPS enabled device (such as cellphone)</summary>
        Device = 3,

        /// <summary>Sourced from a contact card</summary>
        Contact = 4,

        /// <summary>Sourced from a resource (such as a conference room)</summary>
        Resource = 5,
    }
}
