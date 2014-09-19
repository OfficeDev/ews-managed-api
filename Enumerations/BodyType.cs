// ---------------------------------------------------------------------------
// <copyright file="BodyType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the BodyType enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the type of body of an item.
    /// </summary>
    public enum BodyType
    {
        /// <summary>
        /// The body is formatted in HTML.
        /// </summary>
        HTML,

        /// <summary>
        /// The body is in plain text.
        /// </summary>
        Text
    }
}
