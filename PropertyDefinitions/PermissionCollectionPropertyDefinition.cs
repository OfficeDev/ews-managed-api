// ---------------------------------------------------------------------------
// <copyright file="PermissionCollectionPropertyDefinition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PermissionCollectionPropertyDefinition class.</summary>
//-----------------------------------------------------------------------

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
