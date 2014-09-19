// ---------------------------------------------------------------------------
// <copyright file="ServiceXmlSerializationException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ServiceXmlSerializationException class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an error that occurs when the XML for a request cannot be serialized.
    /// </summary>
    [Serializable]
    public class ServiceXmlSerializationException : ServiceLocalException
    {
        /// <summary>
        /// ServiceXmlSerializationException Constructor.
        /// </summary>
        public ServiceXmlSerializationException()
            : base()
        {
        }

        /// <summary>
        /// ServiceXmlSerializationException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        public ServiceXmlSerializationException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// ServiceXmlSerializationException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        /// <param name="innerException">Inner exception.</param>
        public ServiceXmlSerializationException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
