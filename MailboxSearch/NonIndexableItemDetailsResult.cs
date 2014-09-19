// ---------------------------------------------------------------------------
// <copyright file="NonIndexableItemDetailsResult.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the NonIndexableItemDetailsResult class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents non indexable item details result.
    /// </summary>
    public sealed class NonIndexableItemDetailsResult
    {
        /// <summary>
        /// Load from xml
        /// </summary>
        /// <param name="reader">The reader</param>
        /// <returns>Non indexable item details result object</returns>
        internal static NonIndexableItemDetailsResult LoadFromXml(EwsServiceXmlReader reader)
        {
            NonIndexableItemDetailsResult nonIndexableItemDetailsResult = new NonIndexableItemDetailsResult();
            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.NonIndexableItemDetailsResult);

            do
            {
                reader.Read();

                if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Items))
                {
                    List<NonIndexableItem> nonIndexableItems = new List<NonIndexableItem>();
                    if (!reader.IsEmptyElement)
                    {
                        do
                        {
                            reader.Read();
                            NonIndexableItem nonIndexableItem = NonIndexableItem.LoadFromXml(reader);
                            if (nonIndexableItem != null)
                            {
                                nonIndexableItems.Add(nonIndexableItem);
                            }
                        }
                        while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.Items));

                        nonIndexableItemDetailsResult.Items = nonIndexableItems.ToArray();
                    }
                }

                if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.FailedMailboxes))
                {
                    nonIndexableItemDetailsResult.FailedMailboxes = FailedSearchMailbox.LoadFailedMailboxesXml(XmlNamespace.Types, reader);
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.NonIndexableItemDetailsResult));

            return nonIndexableItemDetailsResult;
        }

        /// <summary>
        /// Load from json
        /// </summary>
        /// <param name="jsonObject">The json object</param>
        /// <returns>Non indexable item details result object</returns>
        internal static NonIndexableItemDetailsResult LoadFromJson(JsonObject jsonObject)
        {
            NonIndexableItemDetailsResult nonIndexableItemDetailsResult = new NonIndexableItemDetailsResult();

            return nonIndexableItemDetailsResult;
        }

        /// <summary>
        /// Collection of items
        /// </summary>
        public NonIndexableItem[] Items { get; set; }

        /// <summary>
        /// Failed mailboxes
        /// </summary>
        public FailedSearchMailbox[] FailedMailboxes { get; set; }
    }
}
