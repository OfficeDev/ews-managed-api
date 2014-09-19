// ---------------------------------------------------------------------------
// <copyright file="MimeContent.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MimeContent class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Text;

    /// <summary>
    /// Represents the MIME content of an item.
    /// </summary>
    public sealed class MimeContent : MimeContentBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MimeContent"/> class.
        /// </summary>
        public MimeContent()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MimeContent"/> class.
        /// </summary>
        /// <param name="characterSet">The character set of the content.</param>
        /// <param name="content">The content.</param>
        public MimeContent(string characterSet, byte[] content)
        {
            this.CharacterSet = characterSet;
            this.Content = content;
        }

        #region Object method overrides
        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override string ToString()
        {
            if (this.Content == null)
            {
                return string.Empty;
            }
            else
            {
                try
                {
                    // Try to convert to original MIME content using specified charset. If this fails, 
                    // return the Base64 representation of the content.
                    // Note: Encoding.GetString can throw DecoderFallbackException which is a subclass
                    // of ArgumentException.
                    string charSet = string.IsNullOrEmpty(this.CharacterSet)
                                                ? Encoding.UTF8.EncodingName
                                                : this.CharacterSet;
                    Encoding encoding = Encoding.GetEncoding(charSet);
                    return encoding.GetString(this.Content);
                }
                catch (ArgumentException)
                {
                    return Convert.ToBase64String(this.Content);
                }
            }
        }
        #endregion
    }
}
