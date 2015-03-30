/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.ObjectModel;
    using System.Drawing;
    using System.IO;
    using Microsoft.Exchange.WebServices.Data.Enumerations;

    /// <summary>
    /// Represents the results of a GetUserPhoto operation.
    /// </summary>
    public sealed class GetUserPhotoResults
    {
        /// <summary>
        /// Creates a new instance of the <see cref="GetUserPhotoResults"/> class.
        /// </summary>
        internal GetUserPhotoResults()
        {
        }

        /// <summary>
        /// Accessors for the picture data
        /// </summary>
        public byte[] Photo { get; internal set; }

        /// <summary>
        /// Accessors for the Photo EntityTag
        /// </summary>
        public string EntityTag { get; internal set; }

        /// <summary>
        /// Accessors for the ContentType of the photo
        /// </summary>
        public string ContentType { get; internal set; }

        /// <summary>
        /// Accessors for the Expries header tag
        /// </summary>
        public DateTime Expires { get; internal set; }

        /// <summary>
        /// The status of the photo response
        /// </summary>
        public GetUserPhotoStatus Status { get; internal set; }

        /// <summary>
        /// Creates an image from the photo data
        /// </summary>
        /// <returns>The photo data as an Image</returns>
        public Image AsImage()
        {
            if (this.Photo == null || this.Photo.Length == 0)
            {
                throw new InvalidOperationException("Cannot create image when no photo data returned.");
            }

            Image img;
            using (MemoryStream stream = new MemoryStream(this.Photo))
            {
                img = Image.FromStream(stream);
            }

            return img;
        }
    }
}