// ---------------------------------------------------------------------------
// <copyright file="ManagedFolderInformation.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ManagedFolderInformation class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents information for a managed folder.
    /// </summary>
    public sealed class ManagedFolderInformation : ComplexProperty
    {
        private bool? canDelete;
        private bool? canRenameOrMove;
        private bool? mustDisplayComment;
        private bool? hasQuota;
        private bool? isManagedFoldersRoot;
        private string managedFolderId;
        private string comment;
        private int? storageQuota;
        private int? folderSize;
        private string homePage;

        /// <summary>
        /// Initializes a new instance of the <see cref="ManagedFolderInformation"/> class.
        /// </summary>
        internal ManagedFolderInformation()
            : base()
        {
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.CanDelete:
                    this.canDelete = reader.ReadValue<bool>();
                    return true;
                case XmlElementNames.CanRenameOrMove:
                    this.canRenameOrMove = reader.ReadValue<bool>();
                    return true;
                case XmlElementNames.MustDisplayComment:
                    this.mustDisplayComment = reader.ReadValue<bool>();
                    return true;
                case XmlElementNames.HasQuota:
                    this.hasQuota = reader.ReadValue<bool>();
                    return true;
                case XmlElementNames.IsManagedFoldersRoot:
                    this.isManagedFoldersRoot = reader.ReadValue<bool>();
                    return true;
                case XmlElementNames.ManagedFolderId:
                    this.managedFolderId = reader.ReadValue();
                    return true;
                case XmlElementNames.Comment:
                    reader.TryReadValue(ref this.comment);
                    return true;
                case XmlElementNames.StorageQuota:
                    this.storageQuota = reader.ReadValue<int>();
                    return true;
                case XmlElementNames.FolderSize:
                    this.folderSize = reader.ReadValue<int>();
                    return true;
                case XmlElementNames.HomePage:
                    reader.TryReadValue(ref this.homePage);
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service"></param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlElementNames.CanDelete:
                        this.canDelete = jsonProperty.ReadAsBool(key);
                        break;
                    case XmlElementNames.CanRenameOrMove:
                        this.canRenameOrMove = jsonProperty.ReadAsBool(key);
                        break;
                    case XmlElementNames.MustDisplayComment:
                        this.mustDisplayComment = jsonProperty.ReadAsBool(key);
                        break;
                    case XmlElementNames.HasQuota:
                        this.hasQuota = jsonProperty.ReadAsBool(key);
                        break;
                    case XmlElementNames.IsManagedFoldersRoot:
                        this.isManagedFoldersRoot = jsonProperty.ReadAsBool(key);
                        break;
                    case XmlElementNames.ManagedFolderId:
                        this.managedFolderId = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.Comment:
                        string commentValue = jsonProperty.ReadAsString(key);
                        if (commentValue != null)
                        {
                            this.comment = commentValue;
                        }
                        break;
                    case XmlElementNames.StorageQuota:
                        this.storageQuota = jsonProperty.ReadAsInt(key);
                        break;
                    case XmlElementNames.FolderSize:
                        this.folderSize = jsonProperty.ReadAsInt(key);
                        break;
                    case XmlElementNames.HomePage:
                        string homePageValue = jsonProperty.ReadAsString(key);
                        if (homePageValue != null)
                        {
                            this.homePage = homePageValue;
                        }
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Gets a value indicating whether the user can delete objects in the folder.
        /// </summary>
        public bool? CanDelete
        {
            get { return this.canDelete; }
        }

        /// <summary>
        /// Gets a value indicating whether the user can rename or move objects in the folder.
        /// </summary>
        public bool? CanRenameOrMove
        {
            get { return this.canRenameOrMove; }
        }

        /// <summary>
        /// Gets a value indicating whether the client application must display the Comment property to the user.
        /// </summary>
        public bool? MustDisplayComment
        {
            get { return this.mustDisplayComment; }
        }

        /// <summary>
        /// Gets a value indicating whether the folder has a quota.
        /// </summary>
        public bool? HasQuota
        {
            get { return this.hasQuota; }
        }

        /// <summary>
        /// Gets a value indicating whether the folder is the root of the managed folder hierarchy.
        /// </summary>
        public bool? IsManagedFoldersRoot
        {
            get { return this.isManagedFoldersRoot; }
        }

        /// <summary>
        /// Gets the Managed Folder Id of the folder.
        /// </summary>
        public string ManagedFolderId
        {
            get { return this.managedFolderId; }
        }

        /// <summary>
        /// Gets the comment associated with the folder.
        /// </summary>
        public string Comment
        {
            get { return this.comment; }
        }

        /// <summary>
        /// Gets the storage quota of the folder.
        /// </summary>
        public int? StorageQuota
        {
            get { return this.storageQuota; }
        }

        /// <summary>
        /// Gets the size of the folder.
        /// </summary>
        public int? FolderSize
        {
            get { return this.folderSize; }
        }

        /// <summary>
        /// Gets the home page associated with the folder.
        /// </summary>
        public string HomePage
        {
            get { return this.homePage; }
        }
    }
}
