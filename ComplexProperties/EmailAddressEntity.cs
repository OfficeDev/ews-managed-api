// ---------------------------------------------------------------------------
// <copyright file="EmailAddressEntity.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the EmailAddressEntity class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.IO;

    /// <summary>
    /// Represents an EmailAddressEntity object.
    /// </summary>
    public sealed class EmailAddressEntity : ExtractedEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="EmailAddressEntity"/> class.
        /// </summary>
        internal EmailAddressEntity()
            : base()
        {
        }

        /// <summary>
        /// Gets the meeting suggestion Location.
        /// </summary>
        public string EmailAddress { get; internal set; }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.NlgEmailAddress:
                    this.EmailAddress = reader.ReadElementValue();
                    return true;
                
                default:
                    return base.TryReadElementFromXml(reader);
            }
        }
    }
}
