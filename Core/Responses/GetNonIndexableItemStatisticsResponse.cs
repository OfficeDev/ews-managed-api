// ---------------------------------------------------------------------------
// <copyright file="GetNonIndexableItemStatisticsResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetNonIndexableItemStatisticsResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the GetNonIndexableItemStatistics response.
    /// </summary>
    public sealed class GetNonIndexableItemStatisticsResponse : ServiceResponse
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetNonIndexableItemStatisticsResponse"/> class.
        /// </summary>
        internal GetNonIndexableItemStatisticsResponse()
            : base()
        {
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            this.NonIndexableStatistics = NonIndexableItemStatistic.LoadFromXml(reader);
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            throw new NotImplementedException("GetNonIndexableItemStatistics doesn't support JSON.");
        }

        /// <summary>
        /// List of non indexable statistic
        /// </summary>
        public List<NonIndexableItemStatistic> NonIndexableStatistics
        {
            get;
            internal set;
        }
    }
}
