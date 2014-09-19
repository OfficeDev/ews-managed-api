// ---------------------------------------------------------------------------
// <copyright file="ResponseMessageType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ResponseMessageType enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the type of a ResponseMessage object.
    /// </summary>
    public enum ResponseMessageType
    {
        /// <summary>
        /// The ResponseMessage is a reply to the sender of a message.
        /// </summary>
        Reply,

        /// <summary>
        /// The ResponseMessage is a reply to the sender and all the recipients of a message.
        /// </summary>
        ReplyAll,

        /// <summary>
        /// The ResponseMessage is a forward.
        /// </summary>
        Forward
    }
}
