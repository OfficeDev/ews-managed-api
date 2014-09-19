// ---------------------------------------------------------------------------
// <copyright file="ExchangeVersion.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ExchangeVersion enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Defines the each available Exchange release version
    /// </summary>
    public enum ExchangeVersion
    {
        /// <summary>
        /// Microsoft Exchange 2007, Service Pack 1
        /// </summary>
        Exchange2007_SP1 = 0,

        /// <summary>
        /// Microsoft Exchange 2010
        /// </summary>
        Exchange2010 = 1,

        /// <summary>
        /// Microsoft Exchange 2010, Service Pack 1
        /// </summary>
        Exchange2010_SP1 = 2,

        /// <summary>
        /// Microsoft Exchange 2010, Service Pack 2
        /// </summary>
        Exchange2010_SP2 = 3,

        /// <summary>
        /// Microsoft Exchange 2013
        /// </summary>
        Exchange2013 = 4,

        /// <summary>
        /// Microsoft Exchange 2013 SP1
        /// </summary>
        Exchange2013_SP1 = 5,
    }
}
