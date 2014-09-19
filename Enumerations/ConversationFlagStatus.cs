// ---------------------------------------------------------------------------
// <copyright file="ConversationFlagStatus.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ConversationFlagStatus enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the flag status of a Conversation.
    /// </summary>
    public enum ConversationFlagStatus
    {
        /// <summary>
        /// Not Flagged.
        /// </summary>
        NotFlagged,

        /// <summary>
        /// Flagged.
        /// </summary>
        Flagged,

        /// <summary>
        /// Complete.
        /// </summary>
        Complete
    }
}