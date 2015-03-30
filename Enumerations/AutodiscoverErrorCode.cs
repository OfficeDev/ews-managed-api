/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

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