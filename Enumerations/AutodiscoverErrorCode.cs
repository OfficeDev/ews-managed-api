// ---------------------------------------------------------------------------
// <copyright file="AutodiscoverErrorCode.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AutodiscoverErrorCode enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the error codes that can be returned by the Autodiscover service.
    /// </summary>
    public enum AutodiscoverErrorCode
    {
        /// <summary>
        /// There was no Error.
        /// </summary>
        NoError,

        /// <summary>
        /// The caller must follow the e-mail address redirection that was returned by Autodiscover.
        /// </summary>
        RedirectAddress,

        /// <summary>
        /// The caller must follow the URL redirection that was returned by Autodiscover.
        /// </summary>
        RedirectUrl,

        /// <summary>
        /// The user that was passed in the request is invalid.
        /// </summary>
        InvalidUser,

        /// <summary>
        /// The request is invalid.
        /// </summary>
        InvalidRequest,

        /// <summary>
        /// A specified setting is invalid.
        /// </summary>
        InvalidSetting,

        /// <summary>
        /// A specified setting is not available.
        /// </summary>
        SettingIsNotAvailable,

        /// <summary>
        /// The server is too busy to process the request.
        /// </summary>
        ServerBusy,

        /// <summary>
        /// The requested domain is not valid.
        /// </summary>
        InvalidDomain,

        /// <summary>
        /// The organization is not federated.
        /// </summary>
        NotFederated,

        /// <summary>
        /// Internal server error.
        /// </summary>
        InternalServerError,
    }
}
