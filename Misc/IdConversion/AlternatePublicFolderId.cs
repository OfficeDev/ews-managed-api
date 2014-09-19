// ---------------------------------------------------------------------------
// <copyright file="AlternatePublicFolderId.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AlternatePublicFolderId class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents the Id of a public folder expressed in a specific format.
    /// </summary>
    public class AlternatePublicFolderId : AlternateIdBase
    {
        /// <summary>
        /// Name of schema type used for AlternatePublicFolderId element.
        /// </summary>
        internal const string SchemaTypeName = "AlternatePublicFolderIdType";

        /// <summary>
        /// Initializes a new instance of AlternatePublicFolderId.
        /// </summary>
        public AlternatePublicFolderId()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of AlternatePublicFolderId.
        /// </summary>
        /// <param name="format">The format in which the public folder Id is expressed.</param>
        /// <param name="folderId">The Id of the public folder.</param>
        public AlternatePublicFolderId(IdFormat format, string folderId)
            : base(format)
        {
            this.FolderId = folderId;
        }

        /// <summary>
        /// The Id of the public folder.
        /// </summary>
        public string FolderId
        {
            get; set;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.AlternatePublicFolderId;
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            base.WriteAttributesToXml(writer);

            writer.WriteAttributeValue(XmlAttributeNames.FolderId, this.FolderId);
        }

        /// <summary>
        /// Creates a JSON representation of this object..
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        internal override void InternalToJson(JsonObject jsonObject)
        {
            base.InternalToJson(jsonObject);

            jsonObject.Add(XmlAttributeNames.FolderId, this.FolderId);
        }

        /// <summary>
        /// Loads the attributes from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void LoadAttributesFromXml(EwsServiceXmlReader reader)
        {
            base.LoadAttributesFromXml(reader);

            this.FolderId = reader.ReadAttributeValue(XmlAttributeNames.FolderId);
        }

        /// <summary>
        /// Loads the attributes from json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        internal override void LoadAttributesFromJson(JsonObject responseObject)
        {
            base.LoadAttributesFromJson(responseObject);

            this.FolderId = responseObject.ReadAsString(XmlAttributeNames.FolderId);
        }
    }
}
