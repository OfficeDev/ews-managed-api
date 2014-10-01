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
