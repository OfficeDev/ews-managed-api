/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.ObjectModel;
    using System.IO;
    using System.Xml;

    /// <summary>
    /// Represents the response to a GetPeopleInsights operation.
    /// </summary>
    internal sealed class GetPeopleInsightsResponse : ServiceResponse
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetPeopleInsightsResponse"/> class.
        /// </summary>
        public GetPeopleInsightsResponse()
            : base()
        {
            this.People = new Collection<Person>();
        }

        /// <summary>
        /// Gets the People
        /// </summary>
        internal Collection<Person> People { get; private set; }

        /// <summary>
        /// Read Person from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            EwsUtilities.Assert(
                   this.People != null,
                   "GetPeopleInsightsResponse.ReadElementsFromXml",
                   "People is null.");
            
            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.People);
            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        Person item = new Person();
                        item.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.Person);
                        this.People.Add(item);
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.People));
            }
        }
    }
}