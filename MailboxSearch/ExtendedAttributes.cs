// ---------------------------------------------------------------------------
// <copyright file="ExtendedAttributes.cs" company="Microsoft">
//     Copyright (c) Microsoft. All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// <summary>
//      ExtendedAttributes.cs
// </summary>
// ---------------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Class ExtendedAttributes
    /// </summary>
    public sealed class ExtendedAttributes : List<ExtendedAttribute>
    {
    }

    /// <summary>
    /// Class ExtendedAttribute
    /// </summary>
    public sealed class ExtendedAttribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedAttribute"/> class.
        /// </summary>
        public ExtendedAttribute()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedAttribute"/> class.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        public ExtendedAttribute(string name, string value)
        {
            this.Name = name;
            this.Value = value;
        }

        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        /// <value>The name.</value>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the value.
        /// </summary>
        /// <value>The value.</value>
        public string Value { get; set; }
    }
}
