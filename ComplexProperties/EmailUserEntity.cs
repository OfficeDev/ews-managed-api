// ---------------------------------------------------------------------------
// <copyright file="EmailUserEntity.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the EmailUserEntity class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.IO;

    /// <summary>
    /// Represents an EmailUserEntity object.
    /// </summary>
    public sealed class EmailUserEntity : ComplexProperty
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="EmailUserEntity"/> class.
        /// </summary>
        internal EmailUserEntity()
            : base()
        {
            this.Namespace = XmlNamespace.Types;
        }

        /// <summary>
        /// Gets the EmailUser entity Name.
        /// </summary>
        public string Name { get; internal set; }

        /// <summary>
        /// Gets the EmailUser entity UserId.
        /// </summary>
        public string UserId { get; internal set; }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.NlgName:
                    this.Name = reader.ReadElementValue();
                    return true;

                case XmlElementNames.NlgUserId:
                    this.UserId = reader.ReadElementValue();
                    return true;

                default:
                    return base.TryReadElementFromXml(reader);
            }
        }
    }
}
