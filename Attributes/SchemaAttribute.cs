// ---------------------------------------------------------------------------
// <copyright file="SchemaAttribute.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SchemaAttribute attribute.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// The Schema attribute decorates classes that contain EWS schema definitions.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    internal sealed class SchemaAttribute : Attribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SchemaAttribute"/> class.
        /// </summary>
        internal SchemaAttribute()
        {
        }
    }
}
