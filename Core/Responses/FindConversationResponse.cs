// ---------------------------------------------------------------------------
// <copyright file="FindConversationResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the FindConversationResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Xml;

    /// <summary>
    /// Represents the response to a Conversation search operation.
    /// </summary>
    internal sealed class FindConversationResponse : ServiceResponse
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="FindConversationResponse"/> class.
        /// </summary>
        internal FindConversationResponse() : base()
        {
            this.Results = new FindConversationResults();
        }

        /// <summary>
        /// Gets the collection of conversations in results.
        /// </summary>
        internal Collection<Conversation> Conversations
        {
            get
            {
                return this.Results.Conversations;
            }
        }

        /// <summary>
        /// Gets FindConversation results.
        /// </summary>
        /// <returns>FindConversation results.</returns>
        internal FindConversationResults Results { get; private set; }

        /// <summary>
        /// Read Conversations from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            EwsUtilities.Assert(
                   this.Results.Conversations != null,
                   "FindConversationResponse.ReadElementsFromXml",
                   "conversations is null.");

            EwsUtilities.Assert(
                   this.Results.HighlightTerms != null,
                   "FindConversationResponse.ReadElementsFromXml",
                   "highlightTerms is null.");

            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.Conversations);
            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        Conversation item = EwsUtilities.CreateEwsObjectFromXmlElementName<Conversation>(reader.Service, reader.LocalName);

                        if (item == null)
                        {
                            reader.SkipCurrentElement();
                        }
                        else
                        {
                            item.LoadFromXml(
                                        reader,
                                        true, /* clearPropertyBag */
                                        null,
                                        false  /* summaryPropertiesOnly */);

                            this.Results.Conversations.Add(item);
                        }
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.Conversations));
            }

            reader.Read();

            if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.HighlightTerms) &&
                !reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        HighlightTerm term = new HighlightTerm();

                        term.LoadFromXml(
                            reader,
                            XmlNamespace.Types,
                            XmlElementNames.HighlightTerm);

                        this.Results.HighlightTerms.Add(term);
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.HighlightTerms));
            }

            if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.TotalConversationsInView) && !reader.IsEmptyElement)
            {
                this.Results.TotalCount = reader.ReadElementValue<int>();

                reader.Read();
            }

            if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.IndexedOffset) && !reader.IsEmptyElement)
            {
                this.Results.IndexedOffset = reader.ReadElementValue<int>();

                reader.Read();
            }
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            EwsUtilities.Assert(
                   this.Results.Conversations != null,
                   "FindConversationResponse.ReadElementsFromXml",
                   "conversations is null.");

            EwsUtilities.Assert(
                   this.Results.HighlightTerms != null,
                   "FindConversationResponse.ReadElementsFromXml",
                   "highlightTerms is null.");

            foreach (object conversationObject in responseObject.ReadAsArray(XmlElementNames.Conversations))
            {
                JsonObject jsonConversation = conversationObject as JsonObject;

                Conversation item = EwsUtilities.CreateEwsObjectFromXmlElementName<Conversation>(service, XmlElementNames.Conversation);

                if (item != null)
                {
                    item.LoadFromJson(
                        jsonConversation,
                        service,
                        true,
                        null,
                        false);

                    this.Conversations.Add(item);
                }
            }

            Object[] highlightTermObjects = responseObject.ReadAsArray(XmlElementNames.HighlightTerms);
            if (highlightTermObjects != null)
            {
                foreach (object highlightTermObject in highlightTermObjects)
                {
                    JsonObject jsonHighlightTerm = highlightTermObject as JsonObject;
                    HighlightTerm term = new HighlightTerm();

                    term.LoadFromJson(jsonHighlightTerm, service);
                    this.Results.HighlightTerms.Add(term);
                }
            }

            if (responseObject.ContainsKey(XmlElementNames.TotalConversationsInView))
            {
                this.Results.TotalCount = responseObject.ReadAsInt(XmlElementNames.TotalConversationsInView);
            }

            if (responseObject.ContainsKey(XmlElementNames.IndexedOffset))
            {
                this.Results.IndexedOffset = responseObject.ReadAsInt(XmlElementNames.IndexedOffset);
            }
        }
    }
}
