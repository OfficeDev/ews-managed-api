// ---------------------------------------------------------------------------
// <copyright file="VotingInformation.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the VotingInformation class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.ObjectModel;

    /// <summary>
    /// Represents voting information.
    /// </summary>
    public sealed class VotingInformation : ComplexProperty
    {
        private Collection<VotingOptionData> userOptions = new Collection<VotingOptionData>();
        private string votingResponse;

        /// <summary>
        /// Initializes a new instance of the <see cref="VotingInformation"/> class.
        /// </summary>
        internal VotingInformation()
        {
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.UserOptions:
                    if (!reader.IsEmptyElement)
                    {
                        do
                        {
                            reader.Read();

                            if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.VotingOptionData))
                            {
                                VotingOptionData option = new VotingOptionData();
                                option.LoadFromXml(reader, reader.LocalName);
                                this.userOptions.Add(option);
                            }
                        }
                        while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.UserOptions));
                    }
                    return true;
                case XmlElementNames.VotingResponse:
                    this.votingResponse = reader.ReadElementValue<string>();
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlElementNames.UserOptions:
                        Object[] votingOptionObjects = jsonProperty.ReadAsArray(XmlElementNames.UserOptions);
                        if (votingOptionObjects != null)
                        {
                            foreach (object votingOptionObject in votingOptionObjects)
                            {
                                JsonObject jsonObject = votingOptionObject as JsonObject;
                                VotingOptionData option = new VotingOptionData();

                                option.LoadFromJson(jsonObject, service);
                                this.userOptions.Add(option);
                            }
                        }
                        break;
                    case XmlElementNames.VotingResponse:
                        this.votingResponse = jsonProperty.ReadAsString(key);
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Gets the list of user options.
        /// </summary>
        public Collection<VotingOptionData> UserOptions
        {
            get { return this.userOptions; }
        }

        /// <summary>
        /// Gets the voting response.
        /// </summary>
        public string VotingResponse
        {
            get { return this.votingResponse; }
        }
    }
}
