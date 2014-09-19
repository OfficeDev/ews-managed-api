// ---------------------------------------------------------------------------
// <copyright file="FindItemRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the FindItemRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a FindItem request.
    /// </summary>
    /// <typeparam name="TItem">The type of the item.</typeparam>
    internal sealed class FindItemRequest<TItem> : FindRequest<FindItemResponse<TItem>>
        where TItem : Item
    {
        private Grouping groupBy;

        /// <summary>
        /// Initializes a new instance of the <see cref="FindItemRequest&lt;TItem&gt;"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal FindItemRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Gets the group by clause.
        /// </summary>
        /// <returns>The group by clause, null if the request does not have or support grouping.</returns>
        internal override Grouping GetGroupBy()
        {
            return this.GroupBy;
        }
        
        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal override FindItemResponse<TItem> CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new FindItemResponse<TItem>(this.GroupBy != null, this.View.GetPropertySetOrDefault());
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.FindItem;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.FindItemResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.FindItemResponseMessage;
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
        /// Gets or sets the group by.
        /// </summary>
        /// <value>The group by.</value>
        public Grouping GroupBy
        {
            get { return this.groupBy; }
            set { this.groupBy = value; }
        }
    }
}