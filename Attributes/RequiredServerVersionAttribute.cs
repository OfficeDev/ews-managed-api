/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

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