// ---------------------------------------------------------------------------
// <copyright file="MessageDisposition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MessageDisposition enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines how messages are disposed of in CreateItem and UpdateItem operations.
    /// </summary>
    public enum MessageDisposition
    {
        /// <summary>
        /// Messages are saved but not sent.
        /// </summary>
        SaveOnly,

        /// <summary>
        /// Messages are sent and a copy is saved.
        /// </summary>
        SendAndSaveCopy,

        /// <summary>
        /// Messages are sent but no copy is saved.
        /// </summary>
        SendOnly
    }
}
