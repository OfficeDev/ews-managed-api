// ---------------------------------------------------------------------------
// <copyright file="ServiceVersionException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ServiceVersionException class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an error that occurs when a request cannot be handled due to a service version mismatch.
    /// </summary>
    [Serializable]
    public sealed class ServiceVersionException : ServiceLocalException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ServiceVersionException"/> class.
        /// </summary>
        public ServiceVersionException()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ServiceVersionException"/> class.
        /// </summary>
        /// <param name="message">The error message.</param>
        public ServiceVersionException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ServiceVersionException"/> class.
        /// </summary>
        /// <param name="message">The error message.</param>
        /// <param name="innerException">The inner exception.</param>
        public ServiceVersionException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
