// ---------------------------------------------------------------------------
// <copyright file="MoveCopyFolderRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MoveCopyFolderRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an abstract Move/Copy Folder request.
    /// </summary>
    /// <typeparam name="TResponse">The type of the response.</typeparam>
    internal abstract class MoveCopyFolderRequest<TResponse> : MoveCopyRequest<Folder, TResponse>
        where TResponse : ServiceResponse
    {
        private FolderIdWrapperList folderIds = new FolderIdWrapperList();

        /// <summary>
        /// Validates request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParamCollection(this.FolderIds, "FolderIds");
            this.FolderIds.Validate(this.Service.RequestedServerVersion);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MoveCopyFolderRequest&lt;TResponse&gt;"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal MoveCopyFolderRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Writes the ids as XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteIdsToXml(EwsServiceXmlWriter writer)
        {
            this.folderIds.WriteToXml(
                writer,
                XmlNamespace.Messages,
                XmlElementNames.FolderIds);
        }

        /// <summary>
        /// Adds the ids to json.
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        /// <param name="service">The service.</param>
        internal override void AddIdsToJson(JsonObject jsonObject, ExchangeService service)
        {
            if (this.folderIds.Count > 0)
            {
                jsonObject.Add(XmlElementNames.FolderIds, this.folderIds.InternalToJson(service));
            }
        }

        /// <summary>
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of expected response messages.</returns>
        internal override int GetExpectedResponseMessageCount()
        {
            return this.FolderIds.Count;
        }

        /// <summary>
        /// Gets the folder ids.
        /// </summary>
        /// <value>The folder ids.</value>
        internal FolderIdWrapperList FolderIds
        {
            get { return this.folderIds; }
        }
    }
}
