// ---------------------------------------------------------------------------
// <copyright file="FolderWrapper.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the FolderWrapper enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a folder Id provided by a Folder object.
    /// </summary>
    internal class FolderWrapper : AbstractFolderIdWrapper
    {
        /// <summary>
        /// The Folder object providing the Id.
        /// </summary>
        private Folder folder;

        /// <summary>
        /// Initializes a new instance of FolderWrapper.
        /// </summary>
        /// <param name="folder">The Folder object provinding the Id.</param>
        internal FolderWrapper(Folder folder)
        {
            EwsUtilities.Assert(
                folder != null,
                "FolderWrapper.ctor",
                "folder is null");
            EwsUtilities.Assert(
                !folder.IsNew,
                "FolderWrapper.ctor",
                "folder does not have an Id");

            this.folder = folder;
        }

        /// <summary>
        /// Obtains the Folder object associated with the wrapper.
        /// </summary>
        /// <returns>The Folder object associated with the wrapper.</returns>
        public override Folder GetFolder()
        {
            return this.folder;
        }

        /// <summary>
        /// Writes the Id encapsulated in the wrapper to XML.
        /// </summary>
        /// <param name="writer">The writer to write the Id to.</param>
        internal override void WriteToXml(EwsServiceXmlWriter writer)
        {
            this.folder.Id.WriteToXml(writer);
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
            return this.folder.Id.InternalToJson(service);
        }
    }
}
