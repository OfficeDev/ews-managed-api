// ---------------------------------------------------------------------------
// <copyright file="CreateFolderResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the CreateFolderResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the response to an individual folder creation operation.
    /// </summary>
    internal sealed class CreateFolderResponse : ServiceResponse
    {
        private Folder folder;

        /// <summary>
        /// Initializes a new instance of the <see cref="CreateFolderResponse"/> class.
        /// </summary>
        /// <param name="folder">The folder.</param>
        internal CreateFolderResponse(Folder folder)
            : base()
        {
            this.folder = folder;
        }

        /// <summary>
        /// Gets the object instance.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>Folder.</returns>
        private Folder GetObjectInstance(ExchangeService service, string xmlElementName)
        {
            if (this.folder != null)
            {
                return this.folder;
            }
            else
            {
                return EwsUtilities.CreateEwsObjectFromXmlElementName<Folder>(service, xmlElementName);
            }
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
            base.ReadElementsFromJson(responseObject, service);

            List<Folder> folders = new EwsServiceJsonReader(service).ReadServiceObjectsCollectionFromJson<Folder>(
                responseObject,
                XmlElementNames.Folders,
                this.GetObjectInstance,
                false,  /* clearPropertyBag */
                null,   /* requestedPropertySet */
                false); /* summaryPropertiesOnly */

            this.folder = folders[0];
        }

        /// <summary>
        /// Clears the change log of the created folder if the creation succeeded.
        /// </summary>
        internal override void Loaded()
        {
            if (this.Result == ServiceResult.Success)
            {
                this.folder.ClearChangeLog();
            }
        }
    }
}
