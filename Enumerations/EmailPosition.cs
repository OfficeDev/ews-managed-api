// ---------------------------------------------------------------------------
// <copyright file="EmailPosition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the EmailPosition enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the email position of an extracted entity.
    /// </summary>
    public enum EmailPosition
    {
        /// <summary>
        /// The position is in the latest reply.
        /// </summary>
        LatestReply,

        /// <summary>
        /// The position is not in the latest reply.
        /// </summary>
        Other,

        /// <summary>
        /// The position is in the subject.
        /// </summary>
        Subject,

        /// <summary>
        /// The position is in the signature.
        /// </summary>
        Signature,
    }
}
