// ---------------------------------------------------------------------------
// <copyright file="GetUserPhotoResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetUserPhotoResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Net;
    using System.Xml;

    /// <summary>
    /// Represents the response to GetUserPhoto operation.
    /// </summary>
    internal sealed class GetUserPhotoResponse : ServiceResponse
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetUserPhotoResponse"/> class.
        /// </summary>
        internal GetUserPhotoResponse()
        {
            this.Results = new GetUserPhotoResults();
        }

        /// <summary>
        /// Gets GetUserPhoto results.
        /// </summary>
        /// <returns>GetUserPhoto results.</returns>
        internal GetUserPhotoResults Results { get; private set; }

        /// <summary>
        /// Read Photo results from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            bool hasChanged = reader.ReadElementValue<bool>(XmlNamespace.Messages, XmlElementNames.HasChanged);

            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.PictureData);
            byte[] photoData = reader.ReadBase64ElementValue();

            // We only get a content type if we get a photo
            if (photoData.Length > 0)
            {
                this.Results.Photo = photoData;
                this.Results.ContentType = reader.ReadElementValue(XmlNamespace.Messages, XmlElementNames.ContentType);
            }

            if (hasChanged)
            {
                if (this.Results.Photo.Length == 0)
                {
                    this.Results.Status = GetUserPhotoStatus.PhotoOrUserNotFound;
                }
                else
                {
                    this.Results.Status = GetUserPhotoStatus.PhotoReturned;
                }
            }
            else
            {
                this.Results.Status = GetUserPhotoStatus.PhotoUnchanged;
            }
        }

        /// <summary>
        /// Read Photo response headers
        /// </summary>
        /// <param name="responseHeaders">The response header.</param>
        internal override void ReadHeader(WebHeaderCollection responseHeaders)
        {
            // Parse out the ETag, trimming the quotes
            string etag = responseHeaders[HttpResponseHeader.ETag];
            if (etag != null)
            {
                etag = etag.Replace("\"", "");

                if (etag.Length > 0)
                {
                    this.Results.EntityTag = etag;
                }
            }

            // Parse the Expires tag, leaving it in UTC
            string expires = responseHeaders[HttpResponseHeader.Expires];
            if (expires != null && expires.Length > 0)
            {
                this.Results.Expires = DateTime.Parse(expires, null, DateTimeStyles.RoundtripKind);
            }
        }
    }
}