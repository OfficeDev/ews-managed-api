// ---------------------------------------------------------------------------
// <copyright file="RequiredServerVersionAttribute.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the RequiredServerVersion attribute</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// RequiredServerVersionAttribute decorates classes, methods, properties, enum values with the first Exchange version 
    /// in which they appeared.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    internal sealed class RequiredServerVersionAttribute : Attribute
    {
        /// <summary>
        /// Exchange version.
        /// </summary>
        private ExchangeVersion version;

        /// <summary>
        /// Initializes a new instance of the <see cref="RequiredServerVersionAttribute"/> class.
        /// </summary>
        /// <param name="version">The Exchange version.</param>
        internal RequiredServerVersionAttribute(ExchangeVersion version)
            : base()
        {
            this.version = version;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <value>The name of the XML element.</value>
        internal ExchangeVersion Version
        {
            get
            {
                return this.version;
            }
        }
    }
}
