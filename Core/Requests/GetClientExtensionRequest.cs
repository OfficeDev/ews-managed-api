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
    /// Represents a GetClientExtension request.
    /// </summary>
    internal sealed class GetClientExtensionRequest : SimpleServiceRequestBase
    {
        /// <summary>
        /// The list of extension IDs to return. 
        /// </summary>
        private readonly StringList requestedExtensionIds;

        /// <summary>
        /// Whether enabled extension only should be returned. 
        /// </summary>
        private readonly bool shouldReturnEnabledOnly;

        /// <summary>
        /// Whether it's called from admin or user scope. 
        /// </summary>
        private readonly bool isUserScope;

        /// <summary>
        /// The user identity. 
        /// </summary>
        private readonly string userId;

        /// <summary>
        /// The list of org extension IDs which user enabled. 
        /// </summary>
        private readonly StringList userEnabledExtensionIds;

        /// <summary>
        /// The list of org extension IDs which user disabled. 
        /// </summary>
        private readonly StringList userDisabledExtensionIds;

        /// <summary>
        /// Whether it's called for debugging to retrieve org master table xml
        /// </summary>
        private readonly bool isDebug;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetClientExtensionRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="requestedExtensionIds">An array of requested extension IDs to return.</param>
        /// <param name="shouldReturnEnabledOnly">Whether enabled extension only should be returned, e.g. for user's
        /// OWA/Outlook activation scenario.</param>
        /// <param name="isUserScope">Whether it's called from admin or user scope</param>
        /// <param name="userId">Specifies optional (if called with user scope) user identity. This will allow to do proper
        /// filtering in cases where admin installs an extension for specific users only</param>
        /// <param name="userEnabledExtensionIds">Optional list of org extension IDs which user enabled. This is necessary for
        /// proper result filtering on the server end. E.g. if admin installed N extensions but didn't enable them, it does not
        /// make sense to return manifests for those which user never enabled either. Used only when asked
        /// for enabled extension only (activation scenario).</param>
        /// <param name="userDisabledExtensionIds">Optional list of org extension IDs which user disabled. This is necessary for
        /// proper result filtering on the server end. E.g. if admin installed N optional extensions and enabled them, it does
        /// not make sense to retrieve manifests for extensions which user disabled for him or herself. Used only when asked
        /// for enabled extension only (activation scenario).</param>
        /// <param name="isDebug">Whether it's called for debugging to retrieve org master table xml</param>
        internal GetClientExtensionRequest(
            ExchangeService service,
            StringList requestedExtensionIds,
            bool shouldReturnEnabledOnly,
            bool isUserScope,
            string userId,
            StringList userEnabledExtensionIds,
            StringList userDisabledExtensionIds,
            bool isDebug)
                : base(service)
        {
            this.requestedExtensionIds = requestedExtensionIds;
            this.shouldReturnEnabledOnly = shouldReturnEnabledOnly;
            this.isUserScope = isUserScope;
            this.userId = userId;
            this.userEnabledExtensionIds = userEnabledExtensionIds;
            this.userDisabledExtensionIds = userDisabledExtensionIds;
            this.isDebug = isDebug;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetClientExtensionRequest;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (null != this.requestedExtensionIds && this.requestedExtensionIds.Count > 0)
            {
                writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.ClientExtensionRequestedIds);
                requestedExtensionIds.WriteElementsToXml(writer);
                writer.WriteEndElement();
            }
            
            if (this.isUserScope)
            {
                writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.ClientExtensionUserRequest);

                writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionUserIdentity, this.userId);

                if (shouldReturnEnabledOnly)
                {
                    writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionEnabledOnly, this.shouldReturnEnabledOnly);
                }

                if (null != this.userEnabledExtensionIds && this.userEnabledExtensionIds.Count > 0)
                {
                    writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.ClientExtensionUserEnabled);
                    userEnabledExtensionIds.WriteElementsToXml(writer);
                    writer.WriteEndElement();
                }

                if (null != this.userDisabledExtensionIds && this.userDisabledExtensionIds.Count > 0)
                {
                    writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.ClientExtensionUserDisabled);
                    userDisabledExtensionIds.WriteElementsToXml(writer);
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
            }

            if (isDebug)
            {
                writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.ClientExtensionIsDebug, isDebug);
            }
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetClientExtensionResponse;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            GetClientExtensionResponse response = new GetClientExtensionResponse();
            response.LoadFromXml(reader, XmlElementNames.GetClientExtensionResponse);
            return response;
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2013;
        }

        /// <summary>
        /// Executes this request.
        /// </summary>
        /// <returns>Service response.</returns>
        internal GetClientExtensionResponse Execute()
        {
            GetClientExtensionResponse serviceResponse = (GetClientExtensionResponse)this.InternalExecute();
            serviceResponse.ThrowIfNecessary();
            return serviceResponse;
        }
    }
}