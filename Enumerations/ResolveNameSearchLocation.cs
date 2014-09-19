// ---------------------------------------------------------------------------
// <copyright file="ResolveNameSearchLocation.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ResolveNameSearchLocation enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the location where a ResolveName operation searches for contacts.
    /// </summary>
    public enum ResolveNameSearchLocation
    {
        /// <summary>
        /// The name is resolved against the Global Address List.
        /// </summary>
        DirectoryOnly,

        /// <summary>
        /// The name is resolved against the Global Address List and then against the Contacts folder if no match was found.
        /// </summary>
        DirectoryThenContacts,

        /// <summary>
        /// The name is resolved against the Contacts folder.
        /// </summary>
        ContactsOnly,

        /// <summary>
        /// The name is resolved against the Contacts folder and then against the Global Address List if no match was found.
        /// </summary>
        ContactsThenDirectory
    }
}