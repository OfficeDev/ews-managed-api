// ---------------------------------------------------------------------------
// <copyright file="DisableAppRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DisableAppRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;

    /// <summary>
    /// Disable reason type
    /// </summary>
    public enum DisableReasonType
    {
        /// <summary>
        /// Extension is being disabled with no reason
        /// </summary>
        NoReason,

        /// <summary>
        /// Extension is being disabled from Outlook due to performance reasons
        /// </summary>
        OutlookClientPerformance,

        /// <summary>
        /// Extension is being disabled from OWA due to performance reasons
        /// </summary>
        OWAClientPerformance,

        /// <summary>
        /// Extension is being disabled from MOWA due to performance reasons
        /// </summary>
        MobileClientPerformance
    }

    /// <summary>
    /// Represents a DisableApp request.
    /// </summary>
    internal sealed class DisableAppRequest : SimpleServiceRequestBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DisableAppRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="id">Extension id.</param>
        /// <param name="disableReason">Disable reason.</param>
        internal DisableAppRequest(ExchangeService service, string id, DisableReasonType disableReason)
            : base(service)
        {
            this.Id = id;
            this.DisableReason = disableReason;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.DisableAppRequest;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.ID, this.Id);
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.DisableReason, this.DisableReason);
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.DisableAppResponse;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            DisableAppResponse response = new DisableAppResponse();
            response.LoadFromXml(reader, XmlElementNames.DisableAppResponse);
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
        internal DisableAppResponse Execute()
        {
            DisableAppResponse serviceResponse = (DisableAppResponse)this.InternalExecute();
            serviceResponse.ThrowIfNecessary();
            return serviceResponse;
        }

        /// <summary>
        /// Extension id
        /// </summary>
        private string Id
        {
            get;
            set;
        }

        /// <summary>
        /// Disable reason
        /// </summary>
        private DisableReasonType DisableReason
        {
            get;
            set;
        }
    }
}