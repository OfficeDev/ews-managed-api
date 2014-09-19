// ---------------------------------------------------------------------------
// <copyright file="DnsException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DnsException class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Dns
{
    using System;
    using System.ComponentModel;

    /// <summary>
    /// Represents an error that occurs when performing a DNS operation.
    /// </summary>
    [Serializable]
    internal class DnsException : Win32Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DnsException"/> class.
        /// </summary>
        /// <param name="errorCode">The error code.</param>
        internal DnsException(int errorCode)
            : base(errorCode)
        {
        }
    }
}
