// ---------------------------------------------------------------------------
// <copyright file="MimeContentUTF8.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MimeContentUTF8 class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Text;

    /// <summary>
    /// Represents the MIME content of an item.
    /// </summary>
    public sealed class MimeContentUTF8 : MimeContentBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MimeContentUTF8"/> class.
        /// </summary>
        public MimeContentUTF8()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MimeContentUTF8"/> class.
        /// </summary>
        /// <param name="content">The content.</param>
        public MimeContentUTF8(byte[] content)
        {
            this.CharacterSet = Encoding.UTF8.BodyName;
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
                    // it should always be UTF8 encoding for MimeContentUTF8
                    return Encoding.UTF8.GetString(this.Content);
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
