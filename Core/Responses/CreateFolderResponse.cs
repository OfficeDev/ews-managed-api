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