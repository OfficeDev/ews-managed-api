// ---------------------------------------------------------------------------
// <copyright file="ConversationId.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

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