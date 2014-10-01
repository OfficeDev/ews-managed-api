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
// <summary>Defines the ConvertIdRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a ConvertId request.
    /// </summary>
    internal sealed class ConvertIdRequest : MultiResponseServiceRequest<ConvertIdResponse>, IJsonSerializable
    {
        private IdFormat destinationFormat = IdFormat.EwsId;
        private List<AlternateIdBase> ids = new List<AlternateIdBase>();

        /// <summary>
        /// Initializes a new instance of the <see cref="ConvertIdRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal ConvertIdRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal override ConvertIdResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new ConvertIdResponse();
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.ConvertIdResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.ConvertIdResponseMessage;
        }

        /// <summary>
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of expected response messages.</returns>
        internal override int GetExpectedResponseMessageCount()
        {
            return this.Ids.Count;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.ConvertId;
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParamCollection(this.Ids, "Ids");
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteAttributeValue(XmlAttributeNames.DestinationFormat, this.DestinationFormat);

            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.SourceIds);

            foreach (AlternateIdBase alternateId in this.Ids)
            {
                alternateId.WriteToXml(writer);
            }

            writer.WriteEndElement(); // SourceIds
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
        /// Gets or sets the destination format.
        /// </summary>
        /// <value>The destination format.</value>
        public IdFormat DestinationFormat
        {
            get { return this.destinationFormat; }
            set { this.destinationFormat = value; }
        }

        /// <summary>
        /// Gets the ids.
        /// </summary>
        /// <value>The ids.</value>
        public List<AlternateIdBase> Ids
        {
            get { return this.ids; }
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
            JsonObject jsonObject = new JsonObject();

            jsonObject.Add(XmlAttributeNames.DestinationFormat, this.DestinationFormat);

            List<object> sourceIds = new List<object>();
            foreach (AlternateIdBase id in this.Ids)
            {
                sourceIds.Add(((IJsonSerializable)id).ToJson(service));
            }
            jsonObject.Add(XmlElementNames.SourceIds, sourceIds.ToArray());

            return jsonObject;
        }
    }
}
