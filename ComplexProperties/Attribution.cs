/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents an attribution of an attributed string
    /// </summary>
    public sealed class Attribution : ComplexProperty
    {
        /// <summary>
        /// Attribution id
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Attribution source
        /// </summary>
        public ItemId SourceId { get; set; }

        /// <summary>
        /// Display name
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// Whether writable
        /// </summary>
        public bool IsWritable { get; set; }

        /// <summary>
        /// Whether a quick contact
        /// </summary>
        public bool IsQuickContact { get; set; }

        /// <summary>
        /// Whether hidden
        /// </summary>
        public bool IsHidden { get; set; }

        /// <summary>
        /// Folder id
        /// </summary>
        public FolderId FolderId { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public Attribution()
            : base()
        {
        }

        /// <summary>
        /// Creates an instance with required values only
        /// </summary>
        /// <param name="id">Attribution id</param>
        /// <param name="sourceId">Source Id</param>
        /// <param name="displayName">Display name</param>
        public Attribution(string id, ItemId sourceId, string displayName)
            : this(id, sourceId, displayName, false, false, false, null)
        {
        }

        /// <summary>
        /// Creates an instance with all values
        /// </summary>
        /// <param name="id">Attribution id</param>
        /// <param name="sourceId">Source Id</param>
        /// <param name="displayName">Display name</param>
        /// <param name="isWritable">Whether writable</param>
        /// <param name="isQuickContact">Wther quick contact</param>
        /// <param name="isHidden">Whether hidden</param>
        /// <param name="folderId">Folder id</param>
        public Attribution(string id, ItemId sourceId, string displayName, bool isWritable, bool isQuickContact, bool isHidden, FolderId folderId)
            : this()
        {
            EwsUtilities.ValidateParam(id, "id");
            EwsUtilities.ValidateParam(displayName, "displayName");

            this.Id = id;
            this.SourceId = sourceId;
            this.DisplayName = displayName;
            this.IsWritable = isWritable;
            this.IsQuickContact = isQuickContact;
            this.IsHidden = isHidden;
            this.FolderId = folderId;
        }

        /// <summary>
        /// Tries to read element from XML
        /// </summary>
        /// <param name="reader">XML reader</param>
        /// <returns>Whether reading succeeded</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.Id:
                    this.Id = reader.ReadElementValue();
                    break;
                case XmlElementNames.SourceId:
                    this.SourceId = new ItemId();
                    this.SourceId.LoadFromXml(reader, reader.LocalName);
                    break;
                case XmlElementNames.DisplayName:
                    this.DisplayName = reader.ReadElementValue();
                    break;
                case XmlElementNames.IsWritable:
                    this.IsWritable = reader.ReadElementValue<bool>();
                    break;
                case XmlElementNames.IsQuickContact:
                    this.IsQuickContact = reader.ReadElementValue<bool>();
                    break;
                case XmlElementNames.IsHidden:
                    this.IsHidden = reader.ReadElementValue<bool>();
                    break;
                case XmlElementNames.FolderId:
                    this.FolderId = new FolderId();
                    this.FolderId.LoadFromXml(reader, reader.LocalName);
                    break;

                default:
                    return base.TryReadElementFromXml(reader);
            }

            return true;
        }
    }
}