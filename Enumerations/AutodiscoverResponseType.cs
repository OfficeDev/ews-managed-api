// ---------------------------------------------------------------------------
// <copyright file="AutodiscoverResponseType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AutodiscoverResponseType class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the types of response the Autodiscover service can return.
    /// </summary>
    internal enum AutodiscoverResponseType
    {
        /// <summary>
        /// The request returned an error.
        /// </summary>
        Error,

        /// <summary>
        /// A URL redirection is necessary.
        /// </summary>
        RedirectUrl,

        /// <summary>
        /// An address redirection is necessary.
        /// </summary>
        RedirectAddress,

        /// <summary>
        /// The request succeeded.
        /// </summary>
        Success
    }
}
