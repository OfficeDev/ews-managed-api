#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

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
