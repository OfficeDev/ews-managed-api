// ---------------------------------------------------------------------------
// <copyright file="GetFolderRequestBase.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetFolderRequestBase class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an abstract GetFolder request.
    /// </summary>
    /// <typeparam name="TResponse">The type of ServiceResponse.</typeparam>
    internal abstract class GetFolderRequestBase<TResponse> : GetRequest<Folder, TResponse>
        where TResponse : ServiceResponse
    {
        private FolderIdWrapperList folderIds = new FolderIdWrapperList();

        /// <summary>
        /// Initializes a new instance of the <see cref="GetFolderRequestBase&lt;TResponse&gt;"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        protected GetFolderRequestBase(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParamCollection(this.FolderIds, "FolderIds");
            this.FolderIds.Validate(this.Service.RequestedServerVersion);
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
        /// Gets the type of the service object this request applies to.
        /// </summary>
        /// <returns>The type of service object the request applies to.</returns>
        internal override ServiceObjectType GetServiceObjectType()
        {
            return ServiceObjectType.Folder;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            base.WriteElementsToXml(writer);

            this.FolderIds.WriteToXml(
                writer,
                XmlNamespace.Messages,
                XmlElementNames.FolderIds);
        }

        /// <summary>
        /// Adds the ids to request.
        /// </summary>
        /// <param name="jsonRequest">The json request.</param>
        /// <param name="service">The service.</param>
        internal override void AddIdsToRequest(JsonObject jsonRequest, ExchangeService service)
        {
            jsonRequest.Add(XmlElementNames.FolderIds, this.FolderIds.InternalToJson(service));
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetFolder;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetFolderResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.GetFolderResponseMessage;
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
        /// Gets the folder ids.
        /// </summary>
        /// <value>The folder ids.</value>
        public FolderIdWrapperList FolderIds
        {
            get { return this.folderIds; }
        }
    }
}
