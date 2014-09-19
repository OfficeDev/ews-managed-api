// ---------------------------------------------------------------------------
// <copyright file="ConvertIdResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ConvertIdResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the response to an individual Id conversion operation.
    /// </summary>
    public sealed class ConvertIdResponse : ServiceResponse
    {
        private AlternateIdBase convertedId;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConvertIdResponse"/> class.
        /// </summary>
        internal ConvertIdResponse()
            : base()
        {
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.AlternateId);

            string alternateIdClass = reader.ReadAttributeValue(XmlNamespace.XmlSchemaInstance, XmlAttributeNames.Type);

            int aliasSeparatorIndex = alternateIdClass.IndexOf(':');

            if (aliasSeparatorIndex > -1)
            {
                alternateIdClass = alternateIdClass.Substring(aliasSeparatorIndex + 1);
            }

            // Alternate Id classes are responsible fro reading the AlternateId end element when necessary
            switch (alternateIdClass)
            {
                case AlternateId.SchemaTypeName:
                    this.convertedId = new AlternateId();
                    break;
                case AlternatePublicFolderId.SchemaTypeName:
                    this.convertedId = new AlternatePublicFolderId();
                    break;
                case AlternatePublicFolderItemId.SchemaTypeName:
                    this.convertedId = new AlternatePublicFolderItemId();
                    break;
                default:
                    EwsUtilities.Assert(
                        false,
                        "ConvertIdResponse.ReadElementsFromXml",
                        string.Format("Unknown alternate Id class: {0}", alternateIdClass));
                    break;
            }

            this.convertedId.LoadAttributesFromXml(reader);

            reader.ReadEndElementIfNecessary(XmlNamespace.Messages, XmlElementNames.AlternateId);
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            string alternateIdClass = responseObject.ReadTypeString();

            switch (alternateIdClass)
            {
                case AlternateId.SchemaTypeName:
                    this.convertedId = new AlternateId();
                    break;
                case AlternatePublicFolderId.SchemaTypeName:
                    this.convertedId = new AlternatePublicFolderId();
                    break;
                case AlternatePublicFolderItemId.SchemaTypeName:
                    this.convertedId = new AlternatePublicFolderItemId();
                    break;
                default:
                    EwsUtilities.Assert(
                        false,
                        "ConvertIdResponse.ReadElementsFromXml",
                        string.Format("Unknown alternate Id class: {0}", alternateIdClass));
                    break;
            }

            this.convertedId.LoadAttributesFromJson(responseObject);
        }

        /// <summary>
        /// Gets the converted Id.
        /// </summary>
        public AlternateIdBase ConvertedId
        {
            get { return this.convertedId; }
        }
    }
}
