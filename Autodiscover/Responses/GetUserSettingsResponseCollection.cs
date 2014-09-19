// ---------------------------------------------------------------------------
// <copyright file="GetUserSettingsResponseCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetUserSettingsResponseCollection class.</summary>
//-----------------------------------------------------------------------

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
