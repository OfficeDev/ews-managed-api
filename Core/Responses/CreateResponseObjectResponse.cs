// ---------------------------------------------------------------------------
// <copyright file="CreateResponseObjectResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the CreateResponseObjectResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;

    /// <summary>
    /// Represents response to generic Create request.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    internal sealed class CreateResponseObjectResponse : CreateItemResponseBase
    {
        /// <summary>
        /// Gets Item instance.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>Item.</returns>
        internal override Item GetObjectInstance(ExchangeService service, string xmlElementName)
        {
            return EwsUtilities.CreateEwsObjectFromXmlElementName<Item>(service, xmlElementName);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CreateResponseObjectResponse"/> class.
        /// </summary>
        internal CreateResponseObjectResponse()
            : base()
        {
        }
    }
}
