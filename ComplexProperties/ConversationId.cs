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
// <summary>Defines the ConversationId class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the Id of a Conversation.
    /// </summary>
    public class ConversationId : ServiceId
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConversationId"/> class.
        /// </summary>
        internal ConversationId()
            : base()
        {
        }

        /// <summary>
        /// Defines an implicit conversion between string and ConversationId.
        /// </summary>
        /// <param name="uniqueId">The unique Id to convert to ConversationId.</param>
        /// <returns>A ConversationId initialized with the specified unique Id.</returns>
        public static implicit operator ConversationId(string uniqueId)
        {
            return new ConversationId(uniqueId);
        }

        /// <summary>
        /// Defines an implicit conversion between ConversationId and String.
        /// </summary>
        /// <param name="conversationId">The conversationId to String.</param>
        /// <returns>A ConversationId initialized with the specified unique Id.</returns>
        public static implicit operator String(ConversationId conversationId)
        {
            if (conversationId == null)
            {
                throw new ArgumentNullException("conversationId");
            }

            if (String.IsNullOrEmpty(conversationId.UniqueId))
            {
                return String.Empty;
            }
            else
            {
                // Ignoring the change key info
                return conversationId.UniqueId;
            }
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.ConversationId;
        }

        internal override string GetJsonTypeName()
        {
            return XmlElementNames.ItemId;
        }

        /// <summary>
        /// Initializes a new instance of ConversationId.
        /// </summary>
        /// <param name="uniqueId">The unique Id used to initialize the <see cref="ConversationId"/>.</param>
        public ConversationId(string uniqueId)
            : base(uniqueId)
        {
        }

        /// <summary>
        /// Gets a string representation of the Conversation Id.
        /// </summary>
        /// <returns>The string representation of the conversation id.</returns>
        public override string ToString()
        {
            // We have ignored the change key portion
            return this.UniqueId;
        }
    }
}