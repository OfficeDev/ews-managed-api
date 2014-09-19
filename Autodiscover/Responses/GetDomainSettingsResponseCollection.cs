// ---------------------------------------------------------------------------
// <copyright file="GetDomainSettingsResponseCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetDomainSettingsResponseCollection class.</summary>
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
    /// Represents a collection of responses to GetDomainSettings
    /// </summary>
    public sealed class GetDomainSettingsResponseCollection : AutodiscoverResponseCollection<GetDomainSettingsResponse>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverResponseCollection&lt;TResponse&gt;"/> class.
        /// </summary>
        internal GetDomainSettingsResponseCollection()
        {
        }

        /// <summary>
        /// Create a response instance.
        /// </summary>
        /// <returns>GetDomainSettingsResponse.</returns>
        internal override GetDomainSettingsResponse CreateResponseInstance()
        {
            return new GetDomainSettingsResponse();
        }

        /// <summary>
        /// Gets the name of the response collection XML element.
        /// </summary>
        /// <returns>Response collection XMl element name.</returns>
        internal override string GetResponseCollectionXmlElementName()
        {
            return XmlElementNames.DomainResponses;
        }

        /// <summary>
        /// Gets the name of the response instance XML element.
        /// </summary>
        /// <returns>Response instance XMl element name.</returns>
        internal override string GetResponseInstanceXmlElementName()
        {
            return XmlElementNames.DomainResponse;
        }
    }
}
