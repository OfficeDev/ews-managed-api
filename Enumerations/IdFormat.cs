// ---------------------------------------------------------------------------
// <copyright file="IdFormat.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the IdFormat enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines supported Id formats in ConvertId operations.
    /// </summary>
    public enum IdFormat
    {
        /// <summary>
        /// The EWS Id format used in Exchange 2007 RTM.
        /// </summary>
        EwsLegacyId,

        /// <summary>
        /// The EWS Id format used in Exchange 2007 SP1 and above.
        /// </summary>
        EwsId,

        /// <summary>
        /// The base64-encoded PR_ENTRYID property.
        /// </summary>
        EntryId,

        /// <summary>
        /// The hexadecimal representation  of the PR_ENTRYID property.
        /// </summary>
        HexEntryId,

        /// <summary>
        /// The Store Id format.
        /// </summary>
        StoreId,

        /// <summary>
        /// The Outlook Web Access Id format.
        /// </summary>
        OwaId
    }
}
