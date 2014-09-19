// ---------------------------------------------------------------------------
// <copyright file="ISelfValidate.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ISelfValidate interface.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a class that can self-validate.
    /// </summary>
    internal interface ISelfValidate
    {
        /// <summary>
        /// Validates this instance.
        /// </summary>
        void Validate();
    }
}
