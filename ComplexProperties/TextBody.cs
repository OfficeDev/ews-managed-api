// ---------------------------------------------------------------------------
// <copyright file="TextBody.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the TextBody class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents the body of a message.
    /// </summary>
    public sealed class TextBody : MessageBody
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TextBody"/> class.
        /// </summary>
        public TextBody()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TextBody"/> class.
        /// </summary>
        /// <param name="text">The text of the message body.</param>
        public TextBody(string text)
            : base(BodyType.Text, text)
        {
        }

        /// <summary>
        /// Defines an implicit conversation between a string and TextBody.
        /// </summary>
        /// <param name="textBody">The string to convert to TextBody, assumed to be HTML.</param>
        /// <returns>A TextBody initialized with the specified string.</returns>
        public static implicit operator TextBody(string textBody)
        {
            return new TextBody(textBody);
        }
    }
}
