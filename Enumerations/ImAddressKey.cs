// ---------------------------------------------------------------------------
// <copyright file="ImAddressKey.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ImAddressKey enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines Instant Messaging address entries for a contact.
    /// </summary>
    public enum ImAddressKey
    {
        /// <summary>
        /// The first Instant Messaging address.
        /// </summary>
        ImAddress1,
        
        /// <summary>
        /// The second Instant Messaging address.
        /// </summary>
        ImAddress2,

        /// <summary>
        /// The third Instant Messaging address.
        /// </summary>
        ImAddress3
    }
}
