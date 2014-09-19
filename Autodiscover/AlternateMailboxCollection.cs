// ---------------------------------------------------------------------------
// <copyright file="AlternateMailboxCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AlternateMailboxCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System.Collections.Generic;
    using System.Xml;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents a user setting that is a collection of alternate mailboxes.
    /// </summary>
    public sealed class AlternateMailboxCollection
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AlternateMailboxCollection"/> class.
        /// </summary>
        internal AlternateMailboxCollection()
        {
            this.Entries = new List<AlternateMailbox>();
        }

        /// <summary>
        /// Loads instance of AlternateMailboxCollection from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>AlternateMailboxCollection</returns>
        internal static AlternateMailboxCollection LoadFromXml(EwsXmlReader reader)
        {
            AlternateMailboxCollection instance = new AlternateMailboxCollection();

            do
            {
                reader.Read();

                if ((reader.NodeType == XmlNodeType.Element) && (reader.LocalName == XmlElementNames.AlternateMailbox))
                {
                    instance.Entries.Add(AlternateMailbox.LoadFromXml(reader));
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.AlternateMailboxes));

            return instance;
        }

        /// <summary>
        /// Gets the collection of alternate mailboxes.
        /// </summary>
        public List<AlternateMailbox> Entries
        {
            get; private set;
        }
    }
}
