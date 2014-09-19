// ---------------------------------------------------------------------------
// <copyright file="SuggestionsResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SuggestionsResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Text;

    /// <summary>
    /// Represents the response to a meeting time suggestion availability request.
    /// </summary>
    internal sealed class SuggestionsResponse : ServiceResponse
    {
        private Collection<Suggestion> daySuggestions = new Collection<Suggestion>();

        /// <summary>
        /// Initializes a new instance of the <see cref="SuggestionsResponse"/> class.
        /// </summary>
        internal SuggestionsResponse()
            : base()
        {
        }

        /// <summary>
        /// Loads the suggested days from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadSuggestedDaysFromXml(EwsServiceXmlReader reader)
        {
            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.SuggestionDayResultArray);

            do
            {
                reader.Read();

                if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.SuggestionDayResult))
                {
                    Suggestion daySuggestion = new Suggestion();

                    daySuggestion.LoadFromXml(reader, reader.LocalName);

                    this.daySuggestions.Add(daySuggestion);
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.SuggestionDayResultArray));
        }

        /// <summary>
        /// Gets a list of suggested days.
        /// </summary>
        internal Collection<Suggestion> Suggestions
        {
            get { return this.daySuggestions; }
        }
    }
}
