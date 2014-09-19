// ---------------------------------------------------------------------------
// <copyright file="ResolveNamesResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ResolveNamesResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the response to a name resolution operation.
    /// </summary>
    internal sealed class ResolveNamesResponse : ServiceResponse
    {
        private NameResolutionCollection resolutions;

        /// <summary>
        /// Initializes a new instance of the <see cref="ResolveNamesResponse"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal ResolveNamesResponse(ExchangeService service)
            : base()
        {
            EwsUtilities.Assert(
                service != null,
                "ResolveNamesResponse.ctor",
                "service is null");

            this.resolutions = new NameResolutionCollection(service);
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            this.Resolutions.LoadFromXml(reader);
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            base.ReadElementsFromJson(responseObject, service);

            this.Resolutions.LoadFromJson(responseObject.ReadAsJsonObject(XmlElementNames.ResolutionSet), service);
        }

        /// <summary>
        /// Override base implementation so that API does not throw when name resolution fails to find a match.
        /// EWS returns an error in this case but the API will just return an empty NameResolutionCollection. 
        /// </summary>
        internal override void InternalThrowIfNecessary()
        {
            if (this.ErrorCode != ServiceError.ErrorNameResolutionNoResults)
            {
                base.InternalThrowIfNecessary();
            }
        }

        /// <summary>
        /// Gets a list of name resolution suggestions.
        /// </summary>
        public NameResolutionCollection Resolutions
        {
            get { return this.resolutions; }
        }
    }
}
