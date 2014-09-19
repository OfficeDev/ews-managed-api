// ---------------------------------------------------------------------------
// <copyright file="ServiceResponseException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ServiceResponseException class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents a remote service exception that has a single response.
    /// </summary>
    [Serializable]
    public class ServiceResponseException : ServiceRemoteException
    {
        /// <summary>
        /// Error details Value keys
        /// </summary>
        private const string ExceptionClassKey = "ExceptionClass";
        private const string ExceptionMessageKey = "ExceptionMessage";
        private const string StackTraceKey = "StackTrace";

        /// <summary>
        /// ServiceResponse when service operation failed remotely.
        /// </summary>
        private ServiceResponse response;

        /// <summary>
        /// Initializes a new instance of the <see cref="ServiceResponseException"/> class.
        /// </summary>
        /// <param name="response">The ServiceResponse when service operation failed remotely.</param>
        internal ServiceResponseException(ServiceResponse response)
        {
            this.response = response;
        }

        /// <summary>
        /// Gets the ServiceResponse for the exception.
        /// </summary>
        public ServiceResponse Response
        {
            get { return this.response; }
        }

        /// <summary>
        /// Gets the service error code.
        /// </summary>
        public ServiceError ErrorCode
        {
            get { return this.response.ErrorCode; }
        }

        /// <summary>
        /// Gets a message that describes the current exception.
        /// </summary>
        /// <returns>The error message that explains the reason for the exception.</returns>
        public override string Message
        {
            get
            {
                // Special case for Internal Server Error. If the server returned
                // stack trace information, include it in the exception message.
                if (this.Response.ErrorCode == ServiceError.ErrorInternalServerError)
                {
                    string exceptionClass;
                    string exceptionMessage;
                    string stackTrace;

                    if (this.Response.ErrorDetails.TryGetValue(ExceptionClassKey, out exceptionClass) &&
                        this.Response.ErrorDetails.TryGetValue(ExceptionMessageKey, out exceptionMessage) &&
                        this.Response.ErrorDetails.TryGetValue(StackTraceKey, out stackTrace))
                    {
                        return string.Format(
                            Strings.ServerErrorAndStackTraceDetails,
                            this.Response.ErrorMessage,
                            exceptionClass,
                            exceptionMessage,
                            stackTrace);
                    }
                }

                return this.Response.ErrorMessage;
            }
        }
    }
}
