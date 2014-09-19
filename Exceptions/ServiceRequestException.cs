// ---------------------------------------------------------------------------
// <copyright file="ServiceRequestException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ServiceRequestException class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an error that occurs when a service operation request fails (e.g. connection error).
    /// </summary>
    [Serializable]
    public class ServiceRequestException : ServiceRemoteException
    {
        /// <summary>
        /// ServiceRequestException Constructor.
        /// </summary>
        public ServiceRequestException()
            : base()
        {
        }

        /// <summary>
        /// ServiceRequestException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        public ServiceRequestException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// ServiceRequestException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        /// <param name="innerException">Inner exception.</param>
        public ServiceRequestException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
