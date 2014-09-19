// ---------------------------------------------------------------------------
// <copyright file="FolderIdWrapper.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the FolderIdWrapper enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a folder Id provided by a FolderId object.
    /// </summary>
    internal class FolderIdWrapper : AbstractFolderIdWrapper
    {
        /// <summary>
        /// The FolderId object providing the Id.
        /// </summary>
        private FolderId folderId;

        /// <summary>
        /// Initializes a new instance of FolderIdWrapper.
        /// </summary>
        /// <param name="folderId">The FolderId object providing the Id.</param>
        internal FolderIdWrapper(FolderId folderId)
        {
            EwsUtilities.Assert(
                folderId != null,
                "FolderIdWrapper.ctor",
                "folderId is null");

            this.folderId = folderId;
        }

        /// <summary>
        /// Writes the Id encapsulated in the wrapper to XML.
        /// </summary>
        /// <param name="writer">The writer to write the Id to.</param>
        internal override void WriteToXml(EwsServiceXmlWriter writer)
        {
            this.folderId.WriteToXml(writer);
        }

        /// <summary>
        /// Validates folderId against specified version.
        /// </summary>
        /// <param name="version">The version.</param>
        internal override void Validate(ExchangeVersion version)
        {
            this.folderId.Validate(version);
        }

        /// <summary>
        /// Creates a JSON representation of this object.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object InternalToJson(ExchangeService service)
        {
            return this.folderId.InternalToJson(service);
        }
    }
}
