// ---------------------------------------------------------------------------
// <copyright file="ServiceLocalException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ServiceLocalException class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an error that occurs when a service operation fails locally (e.g. validation error).
    /// </summary>
    [Serializable]
    public class ServiceLocalException : Exception
    {
        /// <summary>
        /// ServiceLocalException Constructor.
        /// </summary>
        public ServiceLocalException()
            : base()
        {
        }

        /// <summary>
        /// ServiceLocalException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        public ServiceLocalException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// ServiceLocalException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        /// <param name="innerException">Inner exception.</param>
        public ServiceLocalException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
