// ---------------------------------------------------------------------------
// <copyright file="UrlEntity.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the UrlEntity class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.IO;

    /// <summary>
    /// Represents an UrlEntity object.
    /// </summary>
    public sealed class UrlEntity : ExtractedEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="UrlEntity"/> class.
        /// </summary>
        internal UrlEntity()
            : base()
        {
        }

        /// <summary>
        /// Gets the meeting suggestion Location.
        /// </summary>
        public string Url { get; internal set; }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.NlgUrl:
                    this.Url = reader.ReadElementValue();
                    return true;
                
                default:
                    return base.TryReadElementFromXml(reader);
            }
        }
    }
}
