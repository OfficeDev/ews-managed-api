// ---------------------------------------------------------------------------
// <copyright file="BatchServiceResponseException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the BatchServiceResponseException class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a remote service exception that can have multiple service responses.
    /// </summary>
    /// <typeparam name="TResponse">The type of the response.</typeparam>
    [Serializable]
    public abstract class BatchServiceResponseException<TResponse> : ServiceRemoteException
        where TResponse : ServiceResponse
    {
        /// <summary>
        /// The list of responses returned by the web method.
        /// </summary>
        private ServiceResponseCollection<TResponse> responses;

        /// <summary>
        /// Initializes a new instance of MultiServiceResponseException.
        /// </summary>
        /// <param name="serviceResponses">The list of responses to be associated with this exception.</param>
        /// <param name="message">The message that describes the error.</param>
        internal BatchServiceResponseException(
            ServiceResponseCollection<TResponse> serviceResponses,
            string message)
            : base(message)
        {
            EwsUtilities.Assert(
                serviceResponses != null,
                "MultiServiceResponseException.ctor",
                "serviceResponses is null");

            this.responses = serviceResponses;
        }

        /// <summary>
        /// Initializes a new instance of MultiServiceResponseException.
        /// </summary>
        /// <param name="serviceResponses">The list of responses to be associated with this exception.</param>
        /// <param name="message">The message that describes the error.</param>
        /// <param name="innerException">The exception that is the cause of the current exception.</param>
        internal BatchServiceResponseException(
            ServiceResponseCollection<TResponse> serviceResponses,
            string message,
            Exception innerException)
            : base(message, innerException)
        {
            EwsUtilities.Assert(
                serviceResponses != null,
                "MultiServiceResponseException.ctor",
                "serviceResponses is null");

            this.responses = serviceResponses;
        }

        /// <summary>
        /// Gets a list of responses returned by the web method.
        /// </summary>
        public ServiceResponseCollection<TResponse> ServiceResponses
        {
            get { return this.responses; }
        }
    }
}
