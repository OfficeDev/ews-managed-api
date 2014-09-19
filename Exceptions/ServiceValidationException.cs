// ---------------------------------------------------------------------------
// <copyright file="ServiceValidationException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ServiceValidationException class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an error that occurs when a validation check fails.
    /// </summary>
    [Serializable]
    public sealed class ServiceValidationException : ServiceLocalException
    {
        /// <summary>
        /// ServiceValidationException Constructor.
        /// </summary>
        public ServiceValidationException()
            : base()
        {
        }

        /// <summary>
        /// ServiceValidationException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        public ServiceValidationException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// ServiceValidationException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        /// <param name="innerException">Inner exception.</param>
        public ServiceValidationException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
