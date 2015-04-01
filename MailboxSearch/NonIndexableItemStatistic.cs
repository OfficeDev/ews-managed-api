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
    using System.Text;

    /// <summary>
    /// Represents non indexable item statistic.
    /// </summary>
    public sealed class NonIndexableItemStatistic
    {
        /// <summary>
        /// Mailbox legacy DN
        /// </summary>
        public string Mailbox { get; set; }

        /// <summary>
        /// Item count
        /// </summary>
        public long ItemCount { get; set; }

        /// <summary>
        /// Error message
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// Load from xml
        /// </summary>
        /// <param name="reader">The reader</param>
        /// <returns>List of non indexable item statistic object</returns>
        internal static List<NonIndexableItemStatistic> LoadFromXml(EwsServiceXmlReader reader)
        {
            List<NonIndexableItemStatistic> results = new List<NonIndexableItemStatistic>();

            reader.Read();
            if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.NonIndexableItemStatistics))
            {
                do
                {
                    reader.Read();
                    if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.NonIndexableItemStatistic))
                    {
                        string mailbox = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Mailbox);
                        int itemCount = reader.ReadElementValue<int>(XmlNamespace.Types, XmlElementNames.ItemCount);
                        string errorMessage = null;
                        if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.ErrorMessage))
                        {
                            errorMessage = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.ErrorMessage);
                        }

                        results.Add(new NonIndexableItemStatistic { Mailbox = mailbox, ItemCount = itemCount, ErrorMessage = errorMessage });
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.NonIndexableItemStatistics));
            }

            return results;
        }
    }
}