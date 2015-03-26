// ---------------------------------------------------------------------------
// <copyright file="GetUserPhotoResults.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetUserPhotoResults class.</summary>
//-----------------------------------------------------------------------

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
