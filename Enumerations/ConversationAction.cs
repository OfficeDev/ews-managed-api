#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

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