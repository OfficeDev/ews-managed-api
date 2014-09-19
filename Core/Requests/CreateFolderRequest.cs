// ---------------------------------------------------------------------------
// <copyright file="CreateFolderRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the CreateFolderRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a CreateFolder request.
    /// </summary>
    internal sealed class CreateFolderRequest : CreateRequest<Folder, ServiceResponse>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CreateFolderRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal CreateFolderRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParam(this.Folders, "Folders");

            // Validate each folder.
            foreach (Folder folder in this.Folders)
            {
                folder.Validate();
            }
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal override ServiceResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new CreateFolderResponse((Folder)EwsUtilities.GetEnumeratedObjectAt(this.Folders, responseIndex));
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.CreateFolder;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.CreateFolderResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.CreateFolderResponseMessage;
        }

        /// <summary>
        /// Gets the name of the parent folder XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetParentFolderXmlElementName()
        {
            return XmlElementNames.ParentFolderId;
        }

        /// <summary>
        /// Gets the name of the object collection XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetObjectCollectionXmlElementName()
        {
            return XmlElementNames.Folders;
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2007_SP1;
        }

        /// <summary>
        /// Gets or sets the folders.
        /// </summary>
        /// <value>The folders.</value>
        public IEnumerable<Folder> Folders
        {
            get { return this.Objects; }
            set { this.Objects = value; }
        }
    }
}
