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

using System.Xml.Serialization;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the type of an InsightSource object.
    /// </summary>
    public enum InsightSourceType
    {
        /// <summary>
        /// The InsightSourceType represents the insight data source from AAD.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        AAD,

        /// <summary>
        /// The InsightSourceType represents the insight data source from Mailbox.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        Mailbox,

        /// <summary>
        /// The InsightSourceType represents the insight data source from LinkedIn.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        LinkedIn,

        /// <summary>
        /// The InsightSourceType represents the insigt data source from Facebook.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        Facebook,

        /// <summary>
        /// The InsightSourceType represents the insigt data source from Delve.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        Delve,

        /// <summary>
        /// The InsightSourceType represents the insigt data source from Satori.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        Satori,
    }
}