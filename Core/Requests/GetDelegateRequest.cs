// ---------------------------------------------------------------------------
// <copyright file="GetDelegateRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetDelegateRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a GetDelegate request.
    /// </summary>
    internal class GetDelegateRequest : DelegateManagementRequestBase<GetDelegateResponse>
    {
        private List<UserId> userIds = new List<UserId>();
        private bool includePermissions;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetDelegateRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal GetDelegateRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Creates the response.
        /// </summary>
        /// <returns>Service response.</returns>
        internal override GetDelegateResponse CreateResponse()
        {
            return new GetDelegateResponse(true);
        }

        /// <summary>
        /// Writes XML attributes.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <remarks>
        /// Subclass will override if it has XML attributes.
        /// </remarks>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            base.WriteAttributesToXml(writer);

            writer.WriteAttributeValue(XmlAttributeNames.IncludePermissions, this.IncludePermissions);
        }

        /// <summary>
        /// Writes the elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            base.WriteElementsToXml(writer);

            if (this.UserIds.Count > 0)
            {
                writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.UserIds);

                foreach (UserId userId in this.UserIds)
                {
                    userId.WriteToXml(writer, XmlElementNames.UserId);
                }

                writer.WriteEndElement(); // UserIds
            }
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetDelegateResponse;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetDelegate;
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
        /// Gets the user ids.
        /// </summary>
        /// <value>The user ids.</value>
        public List<UserId> UserIds
        {
            get { return this.userIds; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether permissions are included.
        /// </summary>
        public bool IncludePermissions
        {
            get { return this.includePermissions; }
            set { this.includePermissions = value; }
        }
    }
}
