// ---------------------------------------------------------------------------
// <copyright file="AutodiscoverRemoteException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AutodiscoverRemoteException class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents an exception that is thrown when the Autodiscover service returns an error.
    /// </summary>
    [Serializable]
    public class AutodiscoverRemoteException : ServiceRemoteException
    {
        private AutodiscoverError error;

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverRemoteException"/> class.
        /// </summary>
        /// <param name="error">The error.</param>
        public AutodiscoverRemoteException(AutodiscoverError error)
            : base()
        {
            this.error = error;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverRemoteException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="error">The error.</param>
        public AutodiscoverRemoteException(string message, AutodiscoverError error)
            : base(message)
        {
            this.error = error;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverRemoteException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="error">The error.</param>
        /// <param name="innerException">The inner exception.</param>
        public AutodiscoverRemoteException(
            string message,
            AutodiscoverError error,
            Exception innerException)
            : base(message, innerException)
        {
            this.error = error;
        }

        /// <summary>
        /// Gets the error.
        /// </summary>
        /// <value>The error.</value>
        public AutodiscoverError Error
        {
            get { return this.error; }
        }
    }
}