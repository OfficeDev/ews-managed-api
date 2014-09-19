// ---------------------------------------------------------------------------
// <copyright file="AttachableAttribute.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AttachableAttribute attribute.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// The Attachable attribute decorates item classes that can be attached to other items.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    internal sealed class AttachableAttribute : Attribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AttachableAttribute"/> class.
        /// </summary>
        internal AttachableAttribute()
            : base()
        {
        }
    }
}
