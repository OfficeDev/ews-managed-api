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