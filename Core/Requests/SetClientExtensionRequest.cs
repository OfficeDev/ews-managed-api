// ---------------------------------------------------------------------------
// <copyright file="SetClientExtensionRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SetClientExtensionRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a SetClientExtension request.
    /// </summary>
    internal sealed class SetClientExtensionRequest : MultiResponseServiceRequest<ServiceResponse>
    {
        /// <summary>
        /// Set action such as install, uninstall and configure. 
        /// </summary>
        private readonly List<SetClientExtensionAction> actions;

        /// <summary>
        /// Initializes a new instance of the <see cref="SetClientExtensionRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="actions">List of actions to execute.</param>
        internal SetClientExtensionRequest(ExchangeService service, List<SetClientExtensionAction> actions)
            : base(service, ServiceErrorHandling.ThrowOnError)
        {
            this.actions = actions;
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParam(this.actions, "actions");
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal override ServiceResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new ServiceResponse();
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
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of expected response messages.</returns>
        internal override int GetExpectedResponseMessageCount()
        {
            return this.actions.Count;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.SetClientExtensionRequest;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.SetClientExtensionResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.SetClientExtensionResponseMessage;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.SetClientExtensionActions);

            foreach (SetClientExtensionAction action in this.actions)
            {
                action.WriteToXml(writer, XmlElementNames.SetClientExtensionAction);
            }

            writer.WriteEndElement();
        }
    }
}