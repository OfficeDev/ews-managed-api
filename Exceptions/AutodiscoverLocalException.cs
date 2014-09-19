// ---------------------------------------------------------------------------
// <copyright file="AutodiscoverLocalException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AutodiscoverLocalException class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an exception that is thrown when the Autodiscover service could not be contacted.
    /// </summary>
    [Serializable]
    public class AutodiscoverLocalException : ServiceLocalException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverLocalException"/> class.
        /// </summary>
        public AutodiscoverLocalException()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverLocalException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        public AutodiscoverLocalException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverLocalException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="innerException">The inner exception.</param>
        public AutodiscoverLocalException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
