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
    using System;
    using System.Collections.Generic;
    using System.Xml;

    /// <summary>
    /// Represents the response to a Persona search operation.
    /// </summary>
    internal sealed class FindPeopleResponse : ServiceResponse
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="FindPeopleResponse"/> class.
        /// </summary>
        internal FindPeopleResponse()
            : base()
        {
            this.Results = new FindPeopleResults();
            this.Sources = new List<string>();
        }

        /// <summary>
        /// Gets the collection of Personas in results.
        /// </summary>
        internal ICollection<Persona> Personas
        {
            get
            {
                return this.Results.Personas;
            }
        }

        /// <summary>
        /// Gets FindPersona results.
        /// </summary>
        /// <returns>FindPersona results.</returns>
        internal FindPeopleResults Results { get; private set; }

        /// <summary>
        /// The transaction ID for the query
        /// </summary>
        internal string TransactionId { get; private set; }

        /// <summary>
        /// Whether we queried GAL
        /// </summary>
        internal ICollection<string> Sources { get; private set; }

        /// <summary>
        /// Read Personas from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            EwsUtilities.Assert(
                   this.Results.Personas != null,
                   "FindPeopleResponse.ReadElementsFromXml",
                   "Personas is null.");

            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.People);
            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        Persona item = EwsUtilities.CreateEwsObjectFromXmlElementName<Persona>(reader.Service, reader.LocalName);

                        if (item == null)
                        {
                            reader.SkipCurrentElement();
                        }
                        else
                        {
                            // Don't clear propertyBag because all properties have been initialized in the persona constructor.
                            item.LoadFromXml(reader, false, null, false);
                            this.Results.Personas.Add(item);
                        }
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.People));
            }

            reader.Read();

            if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.TotalNumberOfPeopleInView) && !reader.IsEmptyElement)
            {
                this.Results.TotalCount = reader.ReadElementValue<int>();
                reader.Read();
            }

            if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.FirstMatchingRowIndex) && !reader.IsEmptyElement)
            {
                this.Results.FirstMatchingRowIndex = reader.ReadElementValue<int>();
                reader.Read();
            }

            if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.FirstLoadedRowIndex) && !reader.IsEmptyElement)
            {
                this.Results.FirstLoadedRowIndex = reader.ReadElementValue<int>();
                reader.Read();
            }

            if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.FindPeopleTransactionId) && !reader.IsEmptyElement)
            {
                this.TransactionId = reader.ReadElementValue<string>();
                reader.Read();
            }
   
            // Future proof by skipping any additional elements before returning
            while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.FindPeopleResponse))
            {
                reader.Read();
            }
        }
    }
}