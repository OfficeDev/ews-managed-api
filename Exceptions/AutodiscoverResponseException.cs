// ---------------------------------------------------------------------------
// <copyright file="AutodiscoverResponseException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AutodiscoverResponseException class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents an exception from an autodiscover error response.
    /// </summary>
    [Serializable]
    public class AutodiscoverResponseException : ServiceRemoteException
    {
        /// <summary>
        /// Error code when Autodiscover service operation failed remotely.
        /// </summary>
        private AutodiscoverErrorCode errorCode;

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverResponseException"/> class.
        /// </summary>
        /// <param name="errorCode">The error code.</param>
        /// <param name="message">The message.</param>
        internal AutodiscoverResponseException(AutodiscoverErrorCode errorCode, string message)
            : base(message)
        {
            this.errorCode = errorCode;
        }

        /// <summary>
        /// Gets the ErrorCode for the exception.
        /// </summary>
        public AutodiscoverErrorCode ErrorCode
        {
            get { return this.errorCode; }
        }
    }
}
