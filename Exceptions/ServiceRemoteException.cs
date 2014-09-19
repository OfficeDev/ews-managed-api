// ---------------------------------------------------------------------------
// <copyright file="ServiceRemoteException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ServiceRemoteException class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an error that occurs when a service operation fails remotely.
    /// </summary>
    [Serializable]
    public class ServiceRemoteException : Exception
    {
        /// <summary>
        /// ServiceRemoteException Constructor.
        /// </summary>
        public ServiceRemoteException()
            : base()
        {
        }

        /// <summary>
        /// ServiceRemoteException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        public ServiceRemoteException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// ServiceRemoteException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        /// <param name="innerException">Inner exception.</param>
        public ServiceRemoteException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
