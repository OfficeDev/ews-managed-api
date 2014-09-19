// ---------------------------------------------------------------------------
// <copyright file="ConnectingIdType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ConnectingIdType enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the type of Id of a ConnectingId object.
    /// </summary>
    public enum ConnectingIdType
    {
        /// <summary>
        /// The connecting Id is a principal name.
        /// </summary>
        PrincipalName,

        /// <summary>
        /// The Id is an SID.
        /// </summary>
        SID,

        /// <summary>
        /// The Id is an SMTP address.
        /// </summary>
        SmtpAddress
    }
}
