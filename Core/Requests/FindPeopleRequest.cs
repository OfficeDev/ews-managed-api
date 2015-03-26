// ---------------------------------------------------------------------------
// <copyright file="FindPeopleRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the FindPeopleRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents a request of a find persona operation
    /// </summary>
    internal sealed class FindPeopleRequest : SimpleServiceRequestBase
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service">Exchange web service</param>
        internal FindPeopleRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Accessors of the view controlling the number of personas returned.
        /// </summary>
        internal ViewBase View { get; set; }

        /// <summary>
        /// Folder Id accessors
        /// </summary>
        internal FolderId FolderId { get; set; }

        /// <summary>
        /// Search filter accessors
        /// Available search filter classes include SearchFilter.IsEqualTo,
        /// SearchFilter.ContainsSubstring and SearchFilter.SearchFilterCollection. If SearchFilter
        /// is null, no search filters are applied.
        /// </summary>
        internal SearchFilter SearchFilter { get; set; }
        
        /// <summary>
        /// Query string accessors
        /// </summary>
        internal string QueryString { get; set; }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();               
            this.View.InternalValidate(this);
        }

        /// <summary>
        /// Writes XML attributes.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            base.WriteAttributesToXml(writer);
            this.View.WriteAttributesToXml(writer);
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (this.SearchFilter != null)
            {
                // Emit the Restriction element
                writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.Restriction);
                this.SearchFilter.WriteToXml(writer);
                writer.WriteEndElement();
            }

            // Emit the View element
            this.View.WriteToXml(writer, null);

            // Emit the SortOrder
            this.View.WriteOrderByToXml(writer);

            // Emit the ParentFolderId element
            if (this.FolderId != null)
            {
                writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.ParentFolderId);
                this.FolderId.WriteToXml(writer);
                writer.WriteEndElement();
            }

            if (!string.IsNullOrEmpty(this.QueryString))
            {
                // Emit the QueryString element
                writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.QueryString);
                writer.WriteValue(this.QueryString, XmlElementNames.QueryString);
                writer.WriteEndElement();
            }

            if (this.Service.RequestedServerVersion >= this.GetMinimumRequiredServerVersion())
            {
                if (this.View.PropertySet != null)
                {
                    this.View.PropertySet.WriteToXml(writer, ServiceObjectType.Persona);
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
            FindPeopleResponse response = new FindPeopleResponse();
            response.LoadFromXml(reader, XmlElementNames.FindPeopleResponse);
            return response;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.FindPeople;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.FindPeopleResponse;
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2013_SP1;
        }

        /// <summary>
        /// Executes this request.
        /// </summary>
        /// <returns>Service response.</returns>
        internal FindPeopleResponse Execute()
        {
            FindPeopleResponse serviceResponse = (FindPeopleResponse)this.InternalExecute();
            serviceResponse.ThrowIfNecessary();
            return serviceResponse;
        }
    }
}