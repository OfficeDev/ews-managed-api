#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

//-----------------------------------------------------------------------
// <summary>Defines the EwsEnumAttribute attribute</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// EwsEnumAttribute decorates enum values with the name that should be used for the
    /// enumeration value in the schema.
    /// If this is used to decorate an enumeration, be sure to add that enum type to the dictionary in EwsUtilities.cs
    /// </summary>
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    internal sealed class EwsEnumAttribute : Attribute
    {
        /// <summary>
        /// The name for the enum value used in the server protocol
        /// </summary>
        private string schemaName;

        /// <summary>
        /// Initializes a new instance of the <see cref="EwsEnumAttribute"/> class.
        /// </summary>
        /// <param name="schemaName">Thename used in the protocol for the enum.</param>
        internal EwsEnumAttribute(string schemaName)
            : base()
        {
            this.schemaName = schemaName;
        }

        /// <summary>
        /// Gets the name of the name used for the enum in the protocol.
        /// </summary>
        /// <value>The name of the name used for the enum in the protocol.</value>
        internal string SchemaName
        {
            get
            {
                return this.schemaName;
            }
        }
    }
}
