// ---------------------------------------------------------------------------
// <copyright file="IFileAttachmentContentHandler.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the IFileAttachmentContentHandler interface.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;

    /// <summary>
    /// Defines a file attachment content handler. Application can implement IFileAttachmentContentHandler
    /// to provide a stream in which the content of file attachment should be written.
    /// </summary>
    public interface IFileAttachmentContentHandler
    {
        /// <summary>
        /// Provides a stream to which the content of the attachment with the specified Id should be written.
        /// </summary>
        /// <param name="attachmentId">The Id of the attachment that is being loaded.</param>
        /// <returns>A Stream to which the content of the attachment will be written.</returns>
        Stream GetOutputStream(string attachmentId);
    }
}
