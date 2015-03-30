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
    /// Represents permission set property definition.
    /// </summary>
    internal class PermissionSetPropertyDefinition : ComplexPropertyDefinitionBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PermissionSetPropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        internal PermissionSetPropertyDefinition(
            string xmlElementName,
            string uri,
            PropertyDefinitionFlags flags,
            ExchangeVersion version)
            : base(
                xmlElementName,
                uri,
                flags,
                version)
        {
        }

        /// <summary>
        /// Creates the property instance.
        /// </summary>
        /// <param name="owner">The owner.</param>
        /// <returns>ComplexProperty.</returns>
        internal override ComplexProperty CreatePropertyInstance(ServiceObject owner)
        {
            Folder folder = owner as Folder;

            EwsUtilities.Assert(
                folder != null,
                "PermissionCollectionPropertyDefinition.CreatePropertyInstance",
                "The owner parameter is not of type Folder or a derived class.");

            return new FolderPermissionCollection(folder);
        }

        /// <summary>
        /// Gets the property type.
        /// </summary>
        public override Type Type
        {
            get { return typeof(FolderPermissionCollection); }
        }
    }
}