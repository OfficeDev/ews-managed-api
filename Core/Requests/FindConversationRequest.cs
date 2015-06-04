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
    /// Represents a request to a Find Conversation operation
    /// </summary>
    internal sealed class FindConversationRequest : SimpleServiceRequestBase
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