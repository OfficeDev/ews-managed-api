// ---------------------------------------------------------------------------
// <copyright file="ExtractedEntity.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ExtractedEntity class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.IO;

    /// <summary>
    /// Represents an ExtractedEntity object.
    /// </summary>
    public abstract class ExtractedEntity : ComplexProperty
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExtractedEntity"/> class.
        /// </summary>
        internal ExtractedEntity()
            : base()
        {
            this.Namespace = XmlNamespace.Types;
        }

        /// <summary>
        /// Gets the Position.
        /// </summary>
        public EmailPosition Position { get; internal set; }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.NlgEmailPosition:
                    string positionAsString = reader.ReadElementValue();

                    if (!string.IsNullOrEmpty(positionAsString))
                    {
                        this.Position = EwsUtilities.Parse<EmailPosition>(positionAsString);
                    }
                    return true;
                
                default:
                    return base.TryReadElementFromXml(reader);
            }
        }
    }
}
