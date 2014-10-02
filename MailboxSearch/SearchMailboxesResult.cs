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

//-----------------------------------------------------------------------
// <summary>Defines the SearchMailboxesResult class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents search mailbox result.
    /// </summary>
    public sealed class SearchMailboxesResult
    {
        /// <summary>
        /// Load from xml
        /// </summary>
        /// <param name="reader">The reader</param>
        /// <returns>Search result object</returns>
        internal static SearchMailboxesResult LoadFromXml(EwsServiceXmlReader reader)
        {
            SearchMailboxesResult searchResult = new SearchMailboxesResult();
            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.SearchMailboxesResult);

            List<MailboxQuery> searchQueries = new List<MailboxQuery>();
            do
            {
                reader.Read();
                if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.SearchQueries))
                {
                    reader.ReadStartElement(XmlNamespace.Types, XmlElementNames.MailboxQuery);
                    string query = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Query);
                    reader.ReadStartElement(XmlNamespace.Types, XmlElementNames.MailboxSearchScopes);
                    List<MailboxSearchScope> mailboxSearchScopes = new List<MailboxSearchScope>();
                    do
                    {
                        reader.Read();
                        if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.MailboxSearchScope))
                        {
                            string mailbox = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Mailbox);
                            reader.ReadStartElement(XmlNamespace.Types, XmlElementNames.SearchScope);
                            string searchScope = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.SearchScope);
                            reader.ReadEndElement(XmlNamespace.Types, XmlElementNames.MailboxSearchScope);
                            mailboxSearchScopes.Add(new MailboxSearchScope(mailbox, (MailboxSearchLocation)Enum.Parse(typeof(MailboxSearchLocation), searchScope)));
                        }
                    }
                    while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.MailboxSearchScopes));
                    reader.ReadEndElementIfNecessary(XmlNamespace.Types, XmlElementNames.MailboxSearchScopes);
                    searchQueries.Add(new MailboxQuery(query, mailboxSearchScopes.ToArray()));
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.SearchQueries));
            reader.ReadEndElementIfNecessary(XmlNamespace.Types, XmlElementNames.SearchQueries);
            searchResult.SearchQueries = searchQueries.ToArray();

            searchResult.ResultType = (SearchResultType)Enum.Parse(typeof(SearchResultType), reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.ResultType));
            searchResult.ItemCount = int.Parse(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.ItemCount));
            searchResult.Size = ulong.Parse(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Size));
            searchResult.PageItemCount = int.Parse(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.PageItemCount));
            searchResult.PageItemSize = ulong.Parse(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.PageItemSize));

            do
            {
                reader.Read();
                if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.KeywordStats))
                {
                    searchResult.KeywordStats = LoadKeywordStatsXml(reader);
                }

                if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Items))
                {
                    searchResult.PreviewItems = LoadPreviewItemsXml(reader);
                }

                if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.FailedMailboxes))
                {
                    searchResult.FailedMailboxes = FailedSearchMailbox.LoadFailedMailboxesXml(XmlNamespace.Types, reader);
                }

                if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Refiners))
                {
                    List<SearchRefinerItem> refiners = new List<SearchRefinerItem>();
                    do
                    {
                        reader.Read();
                        if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Refiner))
                        {
                            refiners.Add(SearchRefinerItem.LoadFromXml(reader));
                        }
                    }
                    while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.Refiners));
                    if (refiners.Count > 0)
                    {
                        searchResult.Refiners = refiners.ToArray();
                    }
                }

                if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.MailboxStats))
                {
                    List<MailboxStatisticsItem> mailboxStats = new List<MailboxStatisticsItem>();
                    do
                    {
                        reader.Read();
                        if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.MailboxStat))
                        {
                            mailboxStats.Add(MailboxStatisticsItem.LoadFromXml(reader));
                        }
                    }
                    while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.MailboxStats));
                    if (mailboxStats.Count > 0)
                    {
                        searchResult.MailboxStats = mailboxStats.ToArray();
                    }
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.SearchMailboxesResult));

            return searchResult;
        }

        /// <summary>
        /// Load from json
        /// </summary>
        /// <param name="jsonObject">The json object</param>
        /// <returns>Search result object</returns>
        internal static SearchMailboxesResult LoadFromJson(JsonObject jsonObject)
        {
            SearchMailboxesResult searchResult = new SearchMailboxesResult();

            return searchResult;
        }

        /// <summary>
        /// Load keyword stats xml
        /// </summary>
        /// <param name="reader">The reader</param>
        /// <returns>Array of keyword statistics</returns>
        private static KeywordStatisticsSearchResult[] LoadKeywordStatsXml(EwsServiceXmlReader reader)
        {
            List<KeywordStatisticsSearchResult> keywordStats = new List<KeywordStatisticsSearchResult>();

            reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, XmlElementNames.KeywordStats);
            do
            {
                reader.Read();
                if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.KeywordStat))
                {
                    KeywordStatisticsSearchResult keywordStat = new KeywordStatisticsSearchResult();
                    keywordStat.Keyword = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Keyword);
                    keywordStat.ItemHits = int.Parse(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.ItemHits));
                    keywordStat.Size = ulong.Parse(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Size));
                    keywordStats.Add(keywordStat);
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.KeywordStats));

            return keywordStats.Count == 0 ? null : keywordStats.ToArray();
        }

        /// <summary>
        /// Load preview items xml
        /// </summary>
        /// <param name="reader">The reader</param>
        /// <returns>Array of preview items</returns>
        private static SearchPreviewItem[] LoadPreviewItemsXml(EwsServiceXmlReader reader)
        {
            List<SearchPreviewItem> previewItems = new List<SearchPreviewItem>();

            reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, XmlElementNames.Items);
            do
            {
                reader.Read();
                if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.SearchPreviewItem))
                {
                    SearchPreviewItem previewItem = new SearchPreviewItem();
                    do
                    {
                        reader.Read();
                        if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Id))
                        {
                            previewItem.Id = new ItemId();
                            previewItem.Id.ReadAttributesFromXml(reader);
                        }
                        else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.ParentId))
                        {
                            previewItem.ParentId = new ItemId();
                            previewItem.ParentId.ReadAttributesFromXml(reader);
                        }
                        else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Mailbox))
                        {
                            previewItem.Mailbox = new PreviewItemMailbox();
                            previewItem.Mailbox.MailboxId = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.MailboxId);
                            previewItem.Mailbox.PrimarySmtpAddress = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.PrimarySmtpAddress);
                        }
                        else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.UniqueHash))
                        {
                            previewItem.UniqueHash = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.UniqueHash);
                        }
                        else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.SortValue))
                        {
                            previewItem.SortValue = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.SortValue);
                        }
                        else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.OwaLink))
                        {
                            previewItem.OwaLink = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.OwaLink);
                        }
                        else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Sender))
                        {
                            previewItem.Sender = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Sender);
                        }
                        else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.ToRecipients))
                        {
                            previewItem.ToRecipients = GetRecipients(reader, XmlElementNames.ToRecipients);
                        }
                        else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.CcRecipients))
                        {
                            previewItem.CcRecipients = GetRecipients(reader, XmlElementNames.CcRecipients);
                        }
                        else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.BccRecipients))
                        {
                            previewItem.BccRecipients = GetRecipients(reader, XmlElementNames.BccRecipients);
                        }
                        else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.CreatedTime))
                        {
                            previewItem.CreatedTime = DateTime.Parse(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.CreatedTime));
                        }
                        else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.ReceivedTime))
                        {
                            previewItem.ReceivedTime = DateTime.Parse(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.ReceivedTime));
                        }
                        else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.SentTime))
                        {
                            previewItem.SentTime = DateTime.Parse(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.SentTime));
                        }
                        else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Subject))
                        {
                            previewItem.Subject = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Subject);
                        }
                        else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Preview))
                        {
                            previewItem.Preview = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Preview);
                        }
                        else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Size))
                        {
                            previewItem.Size = ulong.Parse(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Size));
                        }
                        else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Importance))
                        {
                            previewItem.Importance = (Importance)Enum.Parse(typeof(Importance), reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Importance));
                        }
                        else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Read))
                        {
                            previewItem.Read = bool.Parse(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Read));
                        }
                        else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.HasAttachment))
                        {
                            previewItem.HasAttachment = bool.Parse(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.HasAttachment));
                        }
                        else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.ExtendedProperties))
                        {
                            previewItem.ExtendedProperties = LoadExtendedPropertiesXml(reader);
                        }
                    }
                    while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.SearchPreviewItem));

                    previewItems.Add(previewItem);
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.Items));

            return previewItems.Count == 0 ? null : previewItems.ToArray();
        }

        /// <summary>
        /// Get collection of recipients
        /// </summary>
        /// <param name="reader">The reader</param>
        /// <param name="elementName">Element name</param>
        /// <returns>Array of recipients</returns>
        private static string[] GetRecipients(EwsServiceXmlReader reader, string elementName)
        {
            List<string> toRecipients = new List<string>();
            do
            {
                if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.SmtpAddress))
                {
                    toRecipients.Add(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.SmtpAddress));
                }

                reader.Read();
            }
            while (!reader.IsEndElement(XmlNamespace.Types, elementName));

            return toRecipients.Count == 0 ? null : toRecipients.ToArray();
        }

        /// <summary>
        /// Load extended properties xml
        /// </summary>
        /// <param name="reader">The reader</param>
        /// <returns>Extended properties collection</returns>
        private static ExtendedPropertyCollection LoadExtendedPropertiesXml(EwsServiceXmlReader reader)
        {
            ExtendedPropertyCollection extendedProperties = new ExtendedPropertyCollection();

            reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, XmlElementNames.ExtendedProperties);
            do
            {
                reader.Read();
                if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.ExtendedProperty))
                {
                    extendedProperties.LoadFromXml(reader, XmlElementNames.ExtendedProperty);
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.ExtendedProperties));

            return extendedProperties.Count == 0 ? null : extendedProperties;
        }

        /// <summary>
        /// Search queries
        /// </summary>
        public MailboxQuery[] SearchQueries { get; set; }

        /// <summary>
        /// Result type
        /// </summary>
        public SearchResultType ResultType { get; set; }

        /// <summary>
        /// Item count
        /// </summary>
        public long ItemCount { get; set; }

        /// <summary>
        /// Total size
        /// </summary>
        [CLSCompliant(false)]
        public ulong Size { get; set; }

        /// <summary>
        /// Page item count
        /// </summary>
        public int PageItemCount { get; set; }

        /// <summary>
        /// Total page item size
        /// </summary>
        [CLSCompliant(false)]
        public ulong PageItemSize { get; set; }

        /// <summary>
        /// Keyword statistics search result
        /// </summary>
        public KeywordStatisticsSearchResult[] KeywordStats { get; set; }

        /// <summary>
        /// Search preview items
        /// </summary>
        public SearchPreviewItem[] PreviewItems { get; set; }

        /// <summary>
        /// Failed mailboxes
        /// </summary>
        public FailedSearchMailbox[] FailedMailboxes { get; set; }

        /// <summary>
        /// Refiners
        /// </summary>
        public SearchRefinerItem[] Refiners { get; set; }

        /// <summary>
        /// Mailbox statistics
        /// </summary>
        public MailboxStatisticsItem[] MailboxStats { get; set; }
    }

    #region Search Refiner

    /// <summary>
    /// Search refiner item
    /// </summary>
    public sealed class SearchRefinerItem
    {
        /// <summary>
        /// Refiner name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Refiner value
        /// </summary>
        public string Value { get; set; }

        /// <summary>
        /// Refiner count
        /// </summary>
        public long Count { get; set; }

        /// <summary>
        /// Refiner token, essentially comprises of an operator (i.e. ':' or '>') plus the refiner value
        /// The caller such as Sharepoint can simply append this to refiner name for query refinement
        /// </summary>
        public string Token { get; set; }

        /// <summary>
        /// Load from xml
        /// </summary>
        /// <param name="reader"></param>
        /// <returns></returns>
        internal static SearchRefinerItem LoadFromXml(EwsServiceXmlReader reader)
        {
            SearchRefinerItem sri = new SearchRefinerItem();
            reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, XmlElementNames.Refiner);
            sri.Name = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Name);
            sri.Value = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Value);
            sri.Count = reader.ReadElementValue<long>(XmlNamespace.Types, XmlElementNames.Count);
            sri.Token = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Token);
            return sri;
        }
    }

    #endregion

    #region Mailbox Statistics

    /// <summary>
    /// Mailbox statistics item
    /// </summary>
    public sealed class MailboxStatisticsItem
    {
        /// <summary>
        /// Mailbox id
        /// </summary>
        public string MailboxId { get; set; }

        /// <summary>
        /// Display name
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// Item count
        /// </summary>
        public long ItemCount { get; set; }

        /// <summary>
        /// Total size
        /// </summary>
        [CLSCompliant(false)]
        public ulong Size { get; set; }

        /// <summary>
        /// Load from xml
        /// </summary>
        /// <param name="reader"></param>
        /// <returns></returns>
        internal static MailboxStatisticsItem LoadFromXml(EwsServiceXmlReader reader)
        {
            MailboxStatisticsItem msi = new MailboxStatisticsItem();
            reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, XmlElementNames.MailboxStat);
            msi.MailboxId = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.MailboxId);
            msi.DisplayName = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.DisplayName);
            msi.ItemCount = reader.ReadElementValue<long>(XmlNamespace.Types, XmlElementNames.ItemCount);
            msi.Size = reader.ReadElementValue<ulong>(XmlNamespace.Types, XmlElementNames.Size);
            return msi;
        }
    }

    #endregion
}
