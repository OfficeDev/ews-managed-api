// ---------------------------------------------------------------------------
// <copyright file="ImpersonatedUserId.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ImpersonatedUserId class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents an impersonated user Id.
    /// </summary>
    public sealed class ImpersonatedUserId
    {
        private ConnectingIdType idType;
        private string id;

        /// <summary>
        /// Initializes a new instance of the <see cref="ImpersonatedUserId"/> class.
        /// </summary>
        public ImpersonatedUserId()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ImpersonatedUserId"/> class.
        /// </summary>
        /// <param name="idType">The type of this Id.</param>
        /// <param name="id">The user Id.</param>
        public ImpersonatedUserId(ConnectingIdType idType, string id)
            : this()
        {
            this.idType = idType;
            this.id = id;
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal void WriteToXml(EwsServiceXmlWriter writer)
        {
            if (string.IsNullOrEmpty(this.id))
            {
                throw new ArgumentException(Strings.IdPropertyMustBeSet);
            }

            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.ExchangeImpersonation);
            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.ConnectingSID);

            // For 2007 SP1, use PrimarySmtpAddress for type SmtpAddress
            string connectingIdTypeLocalName =
                (this.idType == ConnectingIdType.SmtpAddress) && (writer.Service.RequestedServerVersion == ExchangeVersion.Exchange2007_SP1) 
                    ? XmlElementNames.PrimarySmtpAddress 
                    : this.IdType.ToString();

            writer.WriteElementValue(
                XmlNamespace.Types,
                connectingIdTypeLocalName,
                this.id);

            writer.WriteEndElement(); // ConnectingSID
            writer.WriteEndElement(); // ExchangeImpersonation
        }

        /// <summary>
        /// Gets or sets the type of the Id.
        /// </summary>
        public ConnectingIdType IdType
        {
            get { return this.idType; }
            set { this.idType = value; }
        }

        /// <summary>
        /// Gets or sets the user Id.
        /// </summary>
        public string Id
        {
            get { return this.id; }
            set { this.id = value; }
        }
    }
}