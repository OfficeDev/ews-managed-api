// ---------------------------------------------------------------------------
// <copyright file="LogicalOperator.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the LogicalOperator enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines a logical operator as used by search filter collections.
    /// </summary>
    public enum LogicalOperator
    {
        /// <summary>
        /// The AND operator.
        /// </summary>
        And,

        /// <summary>
        /// The OR operator.
        /// </summary>
        Or
    }
}
