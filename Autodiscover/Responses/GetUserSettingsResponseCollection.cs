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
    using System.ComponentModel;
    using System.Text;
    using System.Xml;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents a collection of responses to GetUserSettings
    /// </summary>
    public sealed class GetUserSettingsResponseCollection : AutodiscoverResponseCollection<GetUserSettingsResponse>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverResponseCollection&lt;TResponse&gt;"/> class.
        /// </summary>
        internal GetUserSettingsResponseCollection()
        {
        }

        /// <summary>
        /// Create a response instance.
        /// </summary>
        /// <returns>GetUserSettingsResponse.</returns>
        internal override GetUserSettingsResponse CreateResponseInstance()
        {
            return new GetUserSettingsResponse();
        }

        /// <summary>
        /// Gets the name of the response collection XML element.
        /// </summary>
        /// <returns>Response collection XMl element name.</returns>
        internal override string GetResponseCollectionXmlElementName()
        {
            return XmlElementNames.UserResponses;
        }

        /// <summary>
        /// Gets the name of the response instance XML element.
        /// </summary>
        /// <returns>Response instance XMl element name.</returns>
        internal override string GetResponseInstanceXmlElementName()
        {
            return XmlElementNames.UserResponse;
        }
    }
}