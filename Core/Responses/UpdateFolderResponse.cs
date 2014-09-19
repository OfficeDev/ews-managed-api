// ---------------------------------------------------------------------------
// <copyright file="UpdateFolderResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the UpdateFolderResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents response to UpdateFolder request.
    /// </summary>
    internal sealed class UpdateFolderResponse : ServiceResponse
    {
        private Folder folder;

        /// <summary>
        /// Initializes a new instance of the <see cref="UpdateFolderResponse"/> class.
        /// </summary>
        /// <param name="folder">The folder.</param>
        internal UpdateFolderResponse(Folder folder)
            : base()
        {
            EwsUtilities.Assert(
                folder != null,
                "UpdateFolderResponse.ctor",
                "folder is null");

            this.folder = folder;
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            reader.ReadServiceObjectsCollectionFromXml<Folder>(
                XmlElementNames.Folders,
                this.GetObjectInstance,
                false,  /* clearPropertyBag */
                null,   /* requestedPropertySet */
                false); /* summaryPropertiesOnly */
        }

        /// <summary>
        /// Clears the change log of the updated folder if the update succeeded.
        /// </summary>
        internal override void Loaded()
        {
            if (this.Result == ServiceResult.Success)
            {
                this.folder.ClearChangeLog();
            }
        }

        /// <summary>
        /// Gets Folder instance.
        /// </summary>
        /// <param name="session">The session.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>Folder.</returns>
        private Folder GetObjectInstance(ExchangeService session, string xmlElementName)
        {
            return this.folder;
        }
    }
}
