// ---------------------------------------------------------------------------
// <copyright file="FindConversationRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the FindConversationRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a request to a Find Conversation operation
    /// </summary>
    internal sealed class FindConversationRequest : SimpleServiceRequestBase, IJsonSerializable
    {
        private ViewBase view;
        private FolderIdWrapper folderId;
        private string queryString;
        private bool returnHighlightTerms;
        private MailboxSearchLocation? mailboxScope;

        /// <summary>
        /// </summary>
        /// <param name="service"></param>
        internal FindConversationRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Gets or sets the view controlling the number of conversations returned.
        /// </summary>
        public ViewBase View
        {
            get 
            { 
                return this.view; 
            }

            set 
            { 
                this.view = value;
                if (this.view is SeekToConditionItemView)
                {
                    ((SeekToConditionItemView)this.view).SetServiceObjectType(ServiceObjectType.Conversation);
                }
            }
        }

        /// <summary>
        /// Gets or sets folder id
        /// </summary>
        internal FolderIdWrapper FolderId
        {
            get
            {
                return this.folderId;
            }

            set
            {
                this.folderId = value;
            }
        }

        /// <summary>
        /// Gets or sets the query string for search value.
        /// </summary>
        internal string QueryString
        {
            get
            {
                return this.queryString;
            }

            set
            {
                this.queryString = value;
            }
        }

        /// <summary>
        /// Gets or sets the query string highlight terms.
        /// </summary>
        internal bool ReturnHighlightTerms
        {
            get
            {
                return this.returnHighlightTerms;
            }

            set
            {
                this.returnHighlightTerms = value;
            }
        }

        /// <summary>
        /// Gets or sets the mailbox search location to include in the search.
        /// </summary>
        internal MailboxSearchLocation? MailboxScope
        {
            get
            {
                return this.mailboxScope;
            }

            set
            {
                this.mailboxScope = value;
            }
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            this.view.InternalValidate(this);

            // query string parameter is only valid for Exchange2013 or higher
            //
            if (!String.IsNullOrEmpty(this.queryString) &&
                this.Service.RequestedServerVersion < ExchangeVersion.Exchange2013)
            {
                throw new ServiceVersionException(
                    string.Format(
                        Strings.ParameterIncompatibleWithRequestVersion,
                        "queryString",
                        ExchangeVersion.Exchange2013));
            }

            // ReturnHighlightTerms parameter is only valid for Exchange2013 or higher
            //
            if (this.ReturnHighlightTerms &&
                this.Service.RequestedServerVersion < ExchangeVersion.Exchange2013)
            {
                throw new ServiceVersionException(
                    string.Format(
                        Strings.ParameterIncompatibleWithRequestVersion,
                        "returnHighlightTerms",
                        ExchangeVersion.Exchange2013));
            }

            // SeekToConditionItemView is only valid for Exchange2013 or higher
            //
            if ((this.View is SeekToConditionItemView) &&
                this.Service.RequestedServerVersion < ExchangeVersion.Exchange2013)
            {
                throw new ServiceVersionException(
                    string.Format(
                        Strings.ParameterIncompatibleWithRequestVersion,
                        "SeekToConditionItemView",
                        ExchangeVersion.Exchange2013));
            }

            // MailboxScope is only valid for Exchange2013 or higher
            //
            if (this.MailboxScope.HasValue &&
                this.Service.RequestedServerVersion < ExchangeVersion.Exchange2013)
            {
                throw new ServiceVersionException(
                    string.Format(
                        Strings.ParameterIncompatibleWithRequestVersion,
                        "MailboxScope",
                        ExchangeVersion.Exchange2013));
            }
        }

        /// <summary>
        /// Writes XML attributes.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            this.View.WriteAttributesToXml(writer);
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            // Emit the view element
            //
            this.View.WriteToXml(writer, null);

            // Emit the Sort Order
            //
            this.View.WriteOrderByToXml(writer);

            // Emit the Parent Folder Id
            //
            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.ParentFolderId);
            this.FolderId.WriteToXml(writer);
            writer.WriteEndElement();

            // Emit the MailboxScope flag
            // 
            if (this.MailboxScope.HasValue)
            {
                writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.MailboxScope, this.MailboxScope.Value);
            }

            if (!string.IsNullOrEmpty(this.queryString))
            {
                // Emit the QueryString
                //
                writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.QueryString);

                if (this.ReturnHighlightTerms)
                {
                    writer.WriteAttributeString(XmlAttributeNames.ReturnHighlightTerms, this.ReturnHighlightTerms.ToString().ToLowerInvariant());
                }

                writer.WriteValue(this.queryString, XmlElementNames.QueryString);
                writer.WriteEndElement();
            }

            if (this.Service.RequestedServerVersion >= ExchangeVersion.Exchange2013)
            {
                if (this.View.PropertySet != null)
                {
                    this.View.PropertySet.WriteToXml(writer, ServiceObjectType.Conversation);
                }
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

            jsonRequest.Add("Paging", this.View.WritePagingToJson(service));
            this.View.AddJsonProperties(jsonRequest, service);

            JsonObject jsonTargetFolderId = new JsonObject();
            jsonTargetFolderId.Add(XmlElementNames.BaseFolderId, this.FolderId.InternalToJson(service));
            jsonRequest.Add(XmlElementNames.ParentFolderId, jsonTargetFolderId);

            if (this.MailboxScope.HasValue)
            {
                jsonRequest.Add(XmlElementNames.MailboxScope, this.MailboxScope.Value);
            }

            if (!string.IsNullOrEmpty(this.queryString))
            {
                JsonObject jsonQueryString = new JsonObject();
                jsonQueryString.Add(XmlAttributeNames.Value, this.QueryString);

                if (this.ReturnHighlightTerms)
                {
                    jsonQueryString.Add(XmlAttributeNames.ReturnHighlightTerms, this.ReturnHighlightTerms.ToString().ToLowerInvariant());
                }

                jsonRequest.Add(XmlElementNames.QueryString, jsonQueryString);
            }

            if (this.Service.RequestedServerVersion >= ExchangeVersion.Exchange2013)
            {
                if (this.View.PropertySet != null)
                {
                    this.View.PropertySet.WriteGetShapeToJson(jsonRequest, service, ServiceObjectType.Conversation);
                }
            }
            return jsonRequest;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            FindConversationResponse response = new FindConversationResponse();
            response.LoadFromXml(reader, XmlElementNames.FindConversationResponse);
            return response;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="jsonBody">The json body.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(JsonObject jsonBody)
        {
            FindConversationResponse serviceResponse = new FindConversationResponse();
            serviceResponse.LoadFromJson(jsonBody, this.Service);
            return serviceResponse;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.FindConversation;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.FindConversationResponse;
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2010_SP1;
        }

        /// <summary>
        /// Executes this request.
        /// </summary>
        /// <returns>Service response.</returns>
        internal FindConversationResponse Execute()
        {
            FindConversationResponse serviceResponse = (FindConversationResponse)this.InternalExecute();
            serviceResponse.ThrowIfNecessary();
            return serviceResponse;
        }
    }
}