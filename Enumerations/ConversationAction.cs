// ---------------------------------------------------------------------------
// <copyright file="ConversationAction.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ConversationAction enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines actions applicable to Conversation.
    /// </summary>
    internal enum ConversationActionType
    {
        /// <summary>
        /// Categorizes every current and future message in the conversation
        /// </summary>
        AlwaysCategorize,

        /// <summary>
        /// Deletes every current and future message in the conversation
        /// </summary>
        AlwaysDelete,

        /// <summary>
        /// Moves every current and future message in the conversation
        /// </summary>
        AlwaysMove,

        /// <summary>
        /// Deletes current item in context folder in the conversation
        /// </summary>
        Delete,

        /// <summary>
        /// Moves current item in context folder in the conversation
        /// </summary>
        Move,

        /// <summary>
        /// Copies current item in context folder in the conversation
        /// </summary>
        Copy,

        /// <summary>
        /// Marks current item in context folder in the conversation with
        /// provided read state
        /// </summary>
        SetReadState,

        /// <summary>
        /// Set retention policy.
        /// </summary>
        SetRetentionPolicy,

        /// <summary>
        /// Flag current items in context folder in the conversation with provided flag state.
        /// </summary>
        Flag,
    }
}