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
    /// <summary>
    /// Represents a SyncFolderHierarchy request.
    /// </summary>
    internal class SyncFolderHierarchyRequest : MultiResponseServiceRequest<SyncFolderHierarchyResponse>, IJsonSerializable
    {
        private PropertySet propertySet;
        private FolderId syncFolderId;
        private string syncState;

        /// <summary>
        /// Initializes a new instance of the <see cref="SyncFolderHierarchyRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal SyncFolderHierarchyRequest(ExchangeService service)
            : base(service, ServiceErrorHandling.ThrowOnError)
        {
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal override SyncFolderHierarchyResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new SyncFolderHierarchyResponse(this.PropertySet);
        }

        /// <summary>
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of expected responses.</returns>
        internal override int GetExpectedResponseMessageCount()
        {
            return 1;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.SyncFolderHierarchy;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.SyncFolderHierarchyResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.SyncFolderHierarchyResponseMessage;
        }

        /// <summary>
        /// Validates request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParam(this.PropertySet, "PropertySet");
            if (this.SyncFolderId != null)
            {
                this.SyncFolderId.Validate(this.Service.RequestedServerVersion);
            }

            this.PropertySet.ValidateForRequest(this, false /*summaryPropertiesOnly*/);
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            this.PropertySet.WriteToXml(writer, ServiceObjectType.Folder);

            if (this.SyncFolderId != null)
            {
                writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.SyncFolderId);
                this.SyncFolderId.WriteToXml(writer);
                writer.WriteEndElement();
            }

            writer.WriteElementValue(
                XmlNamespace.Messages,
                XmlElementNames.SyncState,
                this.SyncState);
        }

        /// <summary>
        /// Creates a JSON representation of this object.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        object IJsonSerializable.ToJson(ExchangeService service)
        {
            JsonObject jsonRequest = new JsonObject();

            this.propertySet.WriteGetShapeToJson(jsonRequest, service, ServiceObjectType.Folder);

            if (this.SyncFolderId != null)
            {
                JsonObject jsonSyncFolderId = new JsonObject();
                jsonSyncFolderId.Add(XmlElementNames.BaseFolderId, this.SyncFolderId.InternalToJson(service));
                jsonRequest.Add(XmlElementNames.SyncFolderId, jsonSyncFolderId);
            }

            jsonRequest.Add(XmlElementNames.SyncState, this.SyncState);

            return jsonRequest;
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
        /// Gets or sets the property set.
        /// </summary>
        /// <value>The property set.</value>
        public PropertySet PropertySet
        {
            get { return this.propertySet; }
            set { this.propertySet = value; }
        }

        /// <summary>
        /// Gets or sets the sync folder id.
        /// </summary>
        /// <value>The sync folder id.</value>
        public FolderId SyncFolderId
        {
            get { return this.syncFolderId; }
            set { this.syncFolderId = value; }
        }

        /// <summary>
        /// Gets or sets the state of the sync.
        /// </summary>
        /// <value>The state of the sync.</value>
        public string SyncState
        {
            get { return this.syncState; }
            set { this.syncState = value; }
        }
    }
}