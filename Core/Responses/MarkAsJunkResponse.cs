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

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Definition for MarkAsJunkResponse
    /// </summary>
    public class MarkAsJunkResponse : ServiceResponse
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetItemResponse"/> class.
        /// </summary>
        internal MarkAsJunkResponse() : base()
        {
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            reader.Read();

            if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.MovedItemId))
            {
                this.MovedItemId = new ItemId();
                this.MovedItemId.LoadFromXml(reader, XmlNamespace.Messages, XmlElementNames.MovedItemId);

                reader.ReadEndElementIfNecessary(XmlNamespace.Messages, XmlElementNames.MovedItemId);
            }
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">Json response object</param>
        /// <param name="service">Exchange service</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            base.ReadElementsFromJson(responseObject, service);
            if (responseObject.ContainsKey(XmlElementNames.Token))
            {
                this.MovedItemId = new ItemId();
                this.MovedItemId.LoadFromJson(responseObject.ReadAsJsonObject(XmlElementNames.MovedItemId), service);
            }
        }

        /// <summary>
        /// Gets the moved item id.
        /// </summary>
        public ItemId MovedItemId { get; private set; }
    }
}
