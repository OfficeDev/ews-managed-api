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
    /// Represents a SyncFolderItems request.
    /// </summary>
    internal class SyncFolderItemsRequest : MultiResponseServiceRequest<SyncFolderItemsResponse>, IJsonSerializable
    {
        private PropertySet propertySet;
        private FolderId syncFolderId;
        private SyncFolderItemsScope syncScope;
        private string syncState;
        private ItemIdWrapperList ignoredItemIds = new ItemIdWrapperList();
        private int maxChangesReturned = 100;
        private int numberOfDays = 0;

        /// <summary>
        /// Initializes a new instance of the <see cref="SyncFolderItemsRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal SyncFolderItemsRequest(ExchangeService service)
            : base(service, ServiceErrorHandling.ThrowOnError)
        {
        }

        /// <summary>
        /// Creates service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal override SyncFolderItemsResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new SyncFolderItemsResponse(this.PropertySet);
        }

        /// <summary>
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of expected response messages.</returns>
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
            return XmlElementNames.SyncFolderItems;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.SyncFolderItemsResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.SyncFolderItemsResponseMessage;
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParam(this.PropertySet, "PropertySet");
            EwsUtilities.ValidateParam(this.SyncFolderId, "SyncFolderId");
            this.SyncFolderId.Validate(this.Service.RequestedServerVersion);

            // SyncFolderItemsScope enum was introduced with Exchange2010.  Only
            // value NormalItems is valid with previous server versions.
            if (this.Service.RequestedServerVersion < ExchangeVersion.Exchange2010 &&
                this.syncScope != SyncFolderItemsScope.NormalItems)
            {
                throw new ServiceVersionException(
                    string.Format(
                                  Strings.EnumValueIncompatibleWithRequestVersion,
                                  this.syncScope.ToString(),
                                  this.syncScope.GetType().Name,
                                  ExchangeVersion.Exchange2010));
            }

            // NumberOfDays was introduced with Exchange 2013.
            if (this.Service.RequestedServerVersion < ExchangeVersion.Exchange2013 &&
                this.NumberOfDays != 0)
            {
                throw new ServiceVersionException(
                    string.Format(
                                  Strings.ParameterIncompatibleWithRequestVersion,
                                  "numberOfDays",
                                  ExchangeVersion.Exchange2013));
            }

            // SyncFolderItems can only handle summary properties
            this.PropertySet.ValidateForRequest(this, true /*summaryPropertiesOnly*/);
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            this.PropertySet.WriteToXml(writer, ServiceObjectType.Item);

            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.SyncFolderId);
            this.SyncFolderId.WriteToXml(writer);
            writer.WriteEndElement();

            writer.WriteElementValue(
                XmlNamespace.Messages,
                XmlElementNames.SyncState,
                this.SyncState);

            this.IgnoredItemIds.WriteToXml(
                writer,
                XmlNamespace.Messages,
                XmlElementNames.Ignore);

            writer.WriteElementValue(
                XmlNamespace.Messages,
                XmlElementNames.MaxChangesReturned,
                this.MaxChangesReturned);

            if (this.Service.RequestedServerVersion >= ExchangeVersion.Exchange2010)
            {
                writer.WriteElementValue(
                    XmlNamespace.Messages,
                    XmlElementNames.SyncScope,
                    this.syncScope);
            }

            if (this.NumberOfDays != 0)
            {
                writer.WriteElementValue(
                    XmlNamespace.Messages,
                    XmlElementNames.NumberOfDays,
                    this.numberOfDays);
            }
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

            this.propertySet.WriteGetShapeToJson(jsonRequest, service, ServiceObjectType.Item);

            JsonObject jsonSyncFolderId = new JsonObject();
            jsonSyncFolderId.Add(XmlElementNames.BaseFolderId, this.SyncFolderId.InternalToJson(service));
            jsonRequest.Add(XmlElementNames.SyncFolderId, jsonSyncFolderId);

            jsonRequest.Add(XmlElementNames.SyncState, this.SyncState);

            if (this.IgnoredItemIds.Count > 0)
            {
                jsonRequest.Add(XmlElementNames.Ignore, this.IgnoredItemIds.InternalToJson(service));
            }

            jsonRequest.Add(XmlElementNames.MaxChangesReturned, this.MaxChangesReturned);

            if (this.Service.RequestedServerVersion >= ExchangeVersion.Exchange2010)
            {
                jsonRequest.Add(XmlElementNames.SyncScope, this.SyncScope);
            }

            if (this.Service.RequestedServerVersion >= ExchangeVersion.Exchange2013)
            {
                jsonRequest.Add(XmlElementNames.NumberOfDays, this.NumberOfDays);
            }

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
        /// Gets or sets the scope of the sync.
        /// </summary>
        /// <value>The scope of the sync.</value>
        public SyncFolderItemsScope SyncScope
        {
            get { return this.syncScope; }
            set { this.syncScope = value; }
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

        /// <summary>
        /// Gets the list of ignored item ids.
        /// </summary>
        /// <value>The ignored item ids.</value>
        public ItemIdWrapperList IgnoredItemIds
        {
            get { return this.ignoredItemIds; }
        }

        /// <summary>
        /// Gets or sets the maximum number of changes returned by SyncFolderItems.
        /// Values must be between 1 and 512.
        /// Default is 100.
        /// </summary>
        public int MaxChangesReturned
        {
            get
            {
                return this.maxChangesReturned;
            }

            set
            {
                if (value >= 1 && value <= 512)
                {
                    this.maxChangesReturned = value;
                }
                else
                {
                    throw new ArgumentException(Strings.MaxChangesMustBeBetween1And512);
                }
            }
        }

        /// <summary>
        /// Gets or sets the number of days of content returned by SyncFolderItems.
        /// Zero means return all content.
        /// Default is zero.
        /// </summary>
        public int NumberOfDays
        {
            get
            {
                return this.numberOfDays;
            }

            set
            {
                if (value >= 0)
                {
                    this.numberOfDays = value;
                }
                else
                {
                    throw new ArgumentException(Strings.NumberOfDaysMustBePositive);
                }
            }
        }
    }
}