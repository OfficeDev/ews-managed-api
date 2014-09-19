// ---------------------------------------------------------------------------
// <copyright file="ServerBusyException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ServerBusyException class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    
    /// <summary>
    /// Represents a server busy exception found in a service response.
    /// </summary>
    [Serializable]
    public class ServerBusyException : ServiceResponseException
    {
        private const string BackOffMillisecondsKey = @"BackOffMilliseconds";
        private int backOffMilliseconds;

        /// <summary>
        /// Initializes a new instance of the <see cref="ServerBusyException"/> class.
        /// </summary>
        /// <param name="response">The ServiceResponse when service operation failed remotely.</param>
        public ServerBusyException(ServiceResponse response) 
            : base(response)
        {
            if (response.ErrorDetails != null && response.ErrorDetails.ContainsKey(ServerBusyException.BackOffMillisecondsKey))
            {
                Int32.TryParse(response.ErrorDetails[ServerBusyException.BackOffMillisecondsKey], out this.backOffMilliseconds);
            }
        }
        
        /// <summary>
        /// Suggested number of milliseconds to wait before attempting a request again. If zero, 
        /// there is no suggested backoff time.
        /// </summary>
        public int BackOffMilliseconds
        {
            get
            {
                return this.backOffMilliseconds;
            }
        }
    }
}