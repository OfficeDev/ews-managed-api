// ---------------------------------------------------------------------------
// <copyright file="DeleteAttachmentException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DeleteAttachmentException enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an error that occurs when a call to the DeleteAttachment web method fails.
    /// </summary>
    [Serializable]
    public sealed class DeleteAttachmentException : BatchServiceResponseException<DeleteAttachmentResponse>
    {
        /// <summary>
        /// Initializes a new instance of DeleteAttachmentException.
        /// </summary>
        /// <param name="serviceResponses">The list of responses to be associated with this exception.</param>
        /// <param name="message">The message that describes the error.</param>
        internal DeleteAttachmentException(
            ServiceResponseCollection<DeleteAttachmentResponse> serviceResponses,
            string message)
            : base(serviceResponses, message)
        {
        }

        /// <summary>
        /// Initializes a new instance of DeleteAttachmentException.
        /// </summary>
        /// <param name="serviceResponses">The list of responses to be associated with this exception.</param>
        /// <param name="message">The message that describes the error.</param>
        /// <param name="innerException">The exception that is the cause of the current exception.</param>
        internal DeleteAttachmentException(
            ServiceResponseCollection<DeleteAttachmentResponse> serviceResponses,
            string message,
            Exception innerException)
            : base(serviceResponses, message, innerException)
        {
        }
    }
}
