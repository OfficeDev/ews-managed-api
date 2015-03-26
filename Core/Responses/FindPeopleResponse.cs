// ---------------------------------------------------------------------------
// <copyright file="FindPeopleResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the FindPeopleResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
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
        }
    }
}
