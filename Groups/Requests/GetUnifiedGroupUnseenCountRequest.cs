// ---------------------------------------------------------------------------
// <copyright file="GetUnifiedGroupUnseenCountRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetUnifiedGroupUnseenCountRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data.Groups
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Represents a request to the GetUnifiedGroupUnseenCount operation.
    /// </summary>
    internal sealed class GetUnifiedGroupUnseenCountRequest : SimpleServiceRequestBase, IJsonSerializable
    {
        /// <summary>
        /// The last visited time utc for the group
        /// </summary>
        private readonly DateTime lastVisitedTimeUtc;

        /// <summary>
        /// The identify type associated with the group
        /// </summary>
        private readonly UnifiedGroupIdentityType identityType;

        /// <summary>
        /// The value of identity associated with the group
        /// </summary>
        private readonly string identityValue;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetUnifiedGroupUnseenCountRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="lastVisitedTimeUtc">The last visited time utc for the group</param>
        /// <param name="identityType">The identity type for the group</param>
        /// <param name="value">The value associated with the identify type for the group</param>
        internal GetUnifiedGroupUnseenCountRequest(
            ExchangeService service,
            DateTime lastVisitedTimeUtc,
            UnifiedGroupIdentityType identityType,
            string value) : base(service)
        {
            this.lastVisitedTimeUtc = lastVisitedTimeUtc;
            this.identityType = identityType;
            this.identityValue = value;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetUnifiedGroupUnseenCountResponseMessage;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>    
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetUnifiedGroupUnseenCount;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            GetUnifiedGroupUnseenCountResponse response = new GetUnifiedGroupUnseenCountResponse();
            response.LoadFromXml(reader, GetResponseXmlElementName());
            return response;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            UnifiedGroupIdentity groupIdentity = new UnifiedGroupIdentity(this.identityType, this.identityValue);

            groupIdentity.WriteToXml(writer, XmlElementNames.GroupIdentity);
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.LastVisitedTimeUtc, this.lastVisitedTimeUtc.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ"));
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
        internal GetUnifiedGroupUnseenCountResponse Execute()
        {
            return (GetUnifiedGroupUnseenCountResponse)this.InternalExecute();
        }

        /// <summary>
        /// Creates a JSON representation of this object.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        object IJsonSerializable.ToJson(ExchangeService service)
        {
            JsonObject jsonRequest = new JsonObject();

            UnifiedGroupIdentity groupIdentity = new UnifiedGroupIdentity(this.identityType, this.identityValue);

            jsonRequest.Add(XmlElementNames.GroupIdentity, groupIdentity.InternalToJson(service));
            jsonRequest.Add(XmlElementNames.LastVisitedTimeUtc, this.lastVisitedTimeUtc.ToString());
            return jsonRequest;
        }
    }
}
