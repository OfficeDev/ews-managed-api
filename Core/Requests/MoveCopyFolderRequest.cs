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