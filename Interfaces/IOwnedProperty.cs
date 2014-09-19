// ---------------------------------------------------------------------------
// <copyright file="IOwnedProperty.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the IOwnedProperty interface.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Complex properties that implement that interface are owned by an instance
    /// of EwsObject. For this reason, they also cannot be shared.
    /// </summary>
    internal interface IOwnedProperty
    {
        /// <summary>
        /// Gets or sets the owner.
        /// </summary>
        /// <value>The owner.</value>
        ServiceObject Owner { get; set; }
    }
}
