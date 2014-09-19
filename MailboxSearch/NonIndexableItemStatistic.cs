// ---------------------------------------------------------------------------
// <copyright file="NonIndexableItemStatistic.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the NonIndexableItemStatistic class.</summary>
//-----------------------------------------------------------------------

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
