// ---------------------------------------------------------------------------
// <copyright file="MoveCopyFolderResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MoveCopyFolderResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;

    /// <summary>
    /// Represents the base response class for individual folder move and copy operations.
    /// </summary>
    public sealed class MoveCopyFolderResponse : ServiceResponse
    {
        private Folder folder;

        /// <summary>
        /// Initializes a new instance of the <see cref="MoveCopyFolderResponse"/> class.
        /// </summary>
        internal MoveCopyFolderResponse()
            : base()
        {
        }

        /// <summary>
        /// Gets Folder instance.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>Folder.</returns>
        private Folder GetObjectInstance(ExchangeService service, string xmlElementName)
        {
            return EwsUtilities.CreateEwsObjectFromXmlElementName<Folder>(service, xmlElementName);
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            List<Folder> folders = reader.ReadServiceObjectsCollectionFromXml<Folder>(
                XmlElementNames.Folders,
                this.GetObjectInstance,
                false,  /* clearPropertyBag */
                null,   /* requestedPropertySet */
                false); /* summaryPropertiesOnly */

            this.folder = folders[0];
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            EwsServiceJsonReader jsonReader = new EwsServiceJsonReader(service);

            List<Folder> folders = jsonReader.ReadServiceObjectsCollectionFromJson<Folder>(
                responseObject,
                XmlElementNames.Folders,
                this.GetObjectInstance,
                false,  /* clearPropertyBag */
                null,   /* requestedPropertySet */
                false); /* summaryPropertiesOnly */

            this.folder = folders[0];
        }

        /// <summary>
        /// Gets the new (moved or copied) folder.
        /// </summary>
        public Folder Folder
        {
            get { return this.folder; }
        }
    }
}
