#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

//-----------------------------------------------------------------------
// <summary>Defines the AlternatePublicFolderItemId class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the Id of a public folder item expressed in a specific format.
    /// </summary>
    public class AlternatePublicFolderItemId : AlternatePublicFolderId
    {
        /// <summary>
        /// Schema type associated with AlternatePublicFolderItemId.
        /// </summary>
        internal new const string SchemaTypeName = "AlternatePublicFolderItemIdType";

        /// <summary>
        /// Item id.
        /// </summary>
        private string itemId;

        /// <summary>
        /// Initializes a new instance of the <see cref="AlternatePublicFolderItemId"/> class.
        /// </summary>
        public AlternatePublicFolderItemId()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AlternatePublicFolderItemId"/> class.
        /// </summary>
        /// <param name="format">The format in which the public folder item Id is expressed.</param>
        /// <param name="folderId">The Id of the parent public folder of the public folder item.</param>
        /// <param name="itemId">The Id of the public folder item.</param>
        public AlternatePublicFolderItemId(
            IdFormat format,
            string folderId,
            string itemId)
            : base(format, folderId)
        {
            this.itemId = itemId;
        }

        /// <summary>
        /// The Id of the public folder item.
        /// </summary>
        public string ItemId
        {
            get { return this.itemId; }
            set { this.itemId = value; }
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.AlternatePublicFolderItemId;
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            base.WriteAttributesToXml(writer);

            writer.WriteAttributeValue(XmlAttributeNames.ItemId, this.ItemId);
        }

        /// <summary>
        /// Creates a JSON representation of this object..
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        internal override void InternalToJson(JsonObject jsonObject)
        {
            base.InternalToJson(jsonObject);

            jsonObject.Add(XmlAttributeNames.ItemId, this.ItemId);
        }

        /// <summary>
        /// Loads the attributes from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void LoadAttributesFromXml(EwsServiceXmlReader reader)
        {
            base.LoadAttributesFromXml(reader);

            this.itemId = reader.ReadAttributeValue(XmlAttributeNames.ItemId);
        }

        /// <summary>
        /// Loads the attributes from json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        internal override void LoadAttributesFromJson(JsonObject responseObject)
        {
            base.LoadAttributesFromJson(responseObject);

            this.itemId = responseObject.ReadAsString(XmlAttributeNames.ItemId);
        }
    }
}
