// ---------------------------------------------------------------------------
// <copyright file="ServiceXmlDeserializationException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ServiceXmlDeserializationException class.</summary>
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
    public sealed class ServiceXmlDeserializationException : ServiceLocalException
    {
        /// <summary>
        /// ServiceXmlDeserializationException Constructor.
        /// </summary>
        public ServiceXmlDeserializationException()
            : base()
        {
        }

        /// <summary>
        /// ServiceXmlDeserializationException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        public ServiceXmlDeserializationException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// ServiceXmlDeserializationException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        /// <param name="innerException">Inner exception.</param>
        public ServiceXmlDeserializationException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
