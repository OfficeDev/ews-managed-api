// ---------------------------------------------------------------------------
// <copyright file="GetFolderResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetFolderResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the response to an individual folder retrieval operation.
    /// </summary>
    public sealed class GetFolderResponse : ServiceResponse
    {
        private Folder folder;
        private PropertySet propertySet;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetFolderResponse"/> class.
        /// </summary>
        /// <param name="folder">The folder.</param>
        /// <param name="propertySet">The property set from the request.</param>
        internal GetFolderResponse(Folder folder, PropertySet propertySet)
            : base()
        {
            this.folder = folder;
            this.propertySet = propertySet;

            EwsUtilities.Assert(
                this.propertySet != null,
                "GetFolderResponse.ctor",
                "PropertySet should not be null");
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
                true,               /* clearPropertyBag */
                this.propertySet,   /* requestedPropertySet */
                false);             /* summaryPropertiesOnly */

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
                true,               /* clearPropertyBag */
                this.propertySet,   /* requestedPropertySet */
                false);             /* summaryPropertiesOnly */

            this.folder = folders[0];
        }

        /// <summary>
        /// Gets the folder instance.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>Folder.</returns>
        private Folder GetObjectInstance(ExchangeService service, string xmlElementName)
        {
            if (this.Folder != null)
            {
                return this.Folder;
            }
            else
            {
                return EwsUtilities.CreateEwsObjectFromXmlElementName<Folder>(service, xmlElementName);
            }
        }

        /// <summary>
        /// Gets the folder that was retrieved.
        /// </summary>
        public Folder Folder
        {
            get { return this.folder; }
        }
    }
}
