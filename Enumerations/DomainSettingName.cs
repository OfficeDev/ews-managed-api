// ---------------------------------------------------------------------------
// <copyright file="DomainSettingName.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DomainSettingName enumeration.</summary>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Text;

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    /// <summary>
    /// Domain setting names.
    /// </summary>
    public enum DomainSettingName
    {
        /// <summary>
        /// The external URL of the Exchange Web Services.
        /// </summary>
        ExternalEwsUrl,

        /// <summary>
        /// The version of the Exchange server hosting the URL of the Exchange Web Services.
        /// </summary>
        ExternalEwsVersion,
    }
}
