// ---------------------------------------------------------------------------
// <copyright file="ServiceResult.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ServiceResult enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the result of a call to an EWS method. Values in this enumeration have to
    /// be ordered from lowest to highest severity.
    /// </summary>
    public enum ServiceResult
    {
        /// <summary>
        /// The call was successful
        /// </summary>
        Success,

        /// <summary>
        /// The call triggered at least one warning
        /// </summary>
        Warning,

        /// <summary>
        /// The call triggered at least one error
        /// </summary>
        Error
    }
}
