// ---------------------------------------------------------------------------
// <copyright file="ServiceJsonDeserializationException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ServiceJsonDeserializationException class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an error that occurs when the XML for a response cannot be deserialized.
    /// </summary>
    [Serializable]
    public sealed class ServiceJsonDeserializationException : ServiceLocalException
    {
        /// <summary>
        /// ServiceJsonDeserializationException Constructor.
        /// </summary>
        public ServiceJsonDeserializationException()
            : base()
        {
        }

        /// <summary>
        /// ServiceJsonDeserializationException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        public ServiceJsonDeserializationException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// ServiceJsonDeserializationException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        /// <param name="innerException">Inner exception.</param>
        public ServiceJsonDeserializationException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
