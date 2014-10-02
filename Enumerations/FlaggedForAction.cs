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
