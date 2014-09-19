// ---------------------------------------------------------------------------
// <copyright file="CreateAttachmentException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the CreateAttachmentException enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an error that occurs when a call to the CreateAttachment web method fails.
    /// </summary>
    [Serializable]
    public sealed class CreateAttachmentException : BatchServiceResponseException<CreateAttachmentResponse>
    {
        /// <summary>
        /// Initializes a new instance of CreateAttachmentException.
        /// </summary>
        /// <param name="serviceResponses">The list of responses to be associated with this exception.</param>
        /// <param name="message">The message that describes the error.</param>
        internal CreateAttachmentException(
            ServiceResponseCollection<CreateAttachmentResponse> serviceResponses,
            string message)
            : base(serviceResponses, message)
        {
        }

        /// <summary>
        /// Initializes a new instance of CreateAttachmentException.
        /// </summary>
        /// <param name="serviceResponses">The list of responses to be associated with this exception.</param>
        /// <param name="message">The message that describes the error.</param>
        /// <param name="innerException">The exception that is the cause of the current exception.</param>
        internal CreateAttachmentException(
            ServiceResponseCollection<CreateAttachmentResponse> serviceResponses,
            string message,
            Exception innerException)
            : base(serviceResponses, message, innerException)
        {
        }
    }
}
