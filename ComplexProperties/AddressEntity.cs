// ---------------------------------------------------------------------------
// <copyright file="AddressEntity.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AddressEntity class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.IO;

    /// <summary>
    /// Represents an AddressEntity object.
    /// </summary>
    public sealed class AddressEntity : ExtractedEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AddressEntity"/> class.
        /// </summary>
        internal AddressEntity()
            : base()
        {
        }

        /// <summary>
        /// Gets the meeting suggestion Location.
        /// </summary>
        public string Address { get; internal set; }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.NlgAddress:
                    this.Address = reader.ReadElementValue();
                    return true;
                
                default:
                    return base.TryReadElementFromXml(reader);
            }
        }
    }
}
