// ---------------------------------------------------------------------------
// <copyright file="VotingOptionData.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the VotingOptionData class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents voting option information.
    /// </summary>
    public sealed class VotingOptionData : ComplexProperty
    {
        private string displayName;
        private SendPrompt sendPrompt;

        /// <summary>
        /// Initializes a new instance of the <see cref="VotingOptionData"/> class.
        /// </summary>
        internal VotingOptionData()
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
                case XmlElementNames.VotingOptionDisplayName:
                    this.displayName = reader.ReadElementValue<string>();
                    return true;
                case XmlElementNames.SendPrompt:
                    this.sendPrompt = reader.ReadElementValue<SendPrompt>();
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
                    case XmlElementNames.VotingOptionDisplayName:
                        this.displayName = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.SendPrompt:
                        this.sendPrompt = jsonProperty.ReadEnumValue<SendPrompt>(key);
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Gets the display name for the voting option.
        /// </summary>
        public string DisplayName
        {
            get { return this.displayName; }
        }

        /// <summary>
        /// Gets the send prompt.
        /// </summary>
        public SendPrompt SendPrompt
        {
            get { return this.sendPrompt; }
        }
    }
}
