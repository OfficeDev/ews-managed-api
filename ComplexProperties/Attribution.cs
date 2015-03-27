// ---------------------------------------------------------------------------
// <copyright file="Attribution.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the Attribution class.</summary>
//-----------------------------------------------------------------------

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
