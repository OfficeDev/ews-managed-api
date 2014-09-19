// ---------------------------------------------------------------------------
// <copyright file="FlaggedForAction.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the FlaggedForAction enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the follow-up actions that may be stamped on a message.
    /// </summary>
    public enum FlaggedForAction
    {
        /// <summary>
        /// The message is flagged with any action.
        /// </summary>
        Any,

        /// <summary>
        /// The recipient is requested to call the sender.
        /// </summary>
        Call,

        /// <summary>
        /// The recipient is requested not to forward the message.
        /// </summary>
        DoNotForward,

        /// <summary>
        /// The recipient is requested to follow up on the message.
        /// </summary>
        FollowUp,

        /// <summary>
        /// The recipient received the message for information.
        /// </summary>
        FYI,

        /// <summary>
        /// The recipient is requested to forward the message.
        /// </summary>
        Forward,

        /// <summary>
        /// The recipient is informed that a response to the message is not required.
        /// </summary>
        NoResponseNecessary,

        /// <summary>
        /// The recipient is requested to read the message.
        /// </summary>
        Read,

        /// <summary>
        /// The recipient is requested to reply to the sender of the message.
        /// </summary>
        Reply,

        /// <summary>
        /// The recipient is requested to reply to everyone the message was sent to.
        /// </summary>
        ReplyToAll,

        /// <summary>
        /// The recipient is requested to review the message.
        /// </summary>
        Review
    }
}
