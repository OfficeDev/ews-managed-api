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
    using System.Collections.Generic;

    /// <summary>
    /// Represents a permission on a folder.
    /// </summary>
    public sealed class FolderPermission : ComplexProperty
    {
        #region Default permissions

        private static LazyMember<Dictionary<FolderPermissionLevel, FolderPermission>> defaultPermissions = new LazyMember<Dictionary<FolderPermissionLevel, FolderPermission>>(
            delegate()
            {
                Dictionary<FolderPermissionLevel, FolderPermission> result = new Dictionary<FolderPermissionLevel, FolderPermission>();

                FolderPermission permission = new FolderPermission();
                permission.canCreateItems = false;
                permission.canCreateSubFolders = false;
                permission.deleteItems = PermissionScope.None;
                permission.editItems = PermissionScope.None;
                permission.isFolderContact = false;
                permission.isFolderOwner = false;
                permission.isFolderVisible = false;
                permission.readItems = FolderPermissionReadAccess.None;

                result.Add(FolderPermissionLevel.None, permission);

                permission = new FolderPermission();
                permission.canCreateItems = true;
                permission.canCreateSubFolders = false;
                permission.deleteItems = PermissionScope.None;
                permission.editItems = PermissionScope.None;
                permission.isFolderContact = false;
                permission.isFolderOwner = false;
                permission.isFolderVisible = true;
                permission.readItems = FolderPermissionReadAccess.None;

                result.Add(FolderPermissionLevel.Contributor, permission);

                permission = new FolderPermission();
                permission.canCreateItems = false;
                permission.canCreateSubFolders = false;
                permission.deleteItems = PermissionScope.None;
                permission.editItems = PermissionScope.None;
                permission.isFolderContact = false;
                permission.isFolderOwner = false;
                permission.isFolderVisible = true;
                permission.readItems = FolderPermissionReadAccess.FullDetails;

                result.Add(FolderPermissionLevel.Reviewer, permission);

                permission = new FolderPermission();
                permission.canCreateItems = true;
                permission.canCreateSubFolders = false;
                permission.deleteItems = PermissionScope.Owned;
                permission.editItems = PermissionScope.None;
                permission.isFolderContact = false;
                permission.isFolderOwner = false;
                permission.isFolderVisible = true;
                permission.readItems = FolderPermissionReadAccess.FullDetails;

                result.Add(FolderPermissionLevel.NoneditingAuthor, permission);

                permission = new FolderPermission();
                permission.canCreateItems = true;
                permission.canCreateSubFolders = false;
                permission.deleteItems = PermissionScope.Owned;
                permission.editItems = PermissionScope.Owned;
                permission.isFolderContact = false;
                permission.isFolderOwner = false;
                permission.isFolderVisible = true;
                permission.readItems = FolderPermissionReadAccess.FullDetails;

                result.Add(FolderPermissionLevel.Author, permission);

                permission = new FolderPermission();
                permission.canCreateItems = true;
                permission.canCreateSubFolders = true;
                permission.deleteItems = PermissionScope.Owned;
                permission.editItems = PermissionScope.Owned;
                permission.isFolderContact = false;
                permission.isFolderOwner = false;
                permission.isFolderVisible = true;
                permission.readItems = FolderPermissionReadAccess.FullDetails;

                result.Add(FolderPermissionLevel.PublishingAuthor, permission);

                permission = new FolderPermission();
                permission.canCreateItems = true;
                permission.canCreateSubFolders = false;
                permission.deleteItems = PermissionScope.All;
                permission.editItems = PermissionScope.All;
                permission.isFolderContact = false;
                permission.isFolderOwner = false;
                permission.isFolderVisible = true;
                permission.readItems = FolderPermissionReadAccess.FullDetails;

                result.Add(FolderPermissionLevel.Editor, permission);

                permission = new FolderPermission();
                permission.canCreateItems = true;
                permission.canCreateSubFolders = true;
                permission.deleteItems = PermissionScope.All;
                permission.editItems = PermissionScope.All;
                permission.isFolderContact = false;
                permission.isFolderOwner = false;
                permission.isFolderVisible = true;
                permission.readItems = FolderPermissionReadAccess.FullDetails;

                result.Add(FolderPermissionLevel.PublishingEditor, permission);

                permission = new FolderPermission();
                permission.canCreateItems = true;
                permission.canCreateSubFolders = true;
                permission.deleteItems = PermissionScope.All;
                permission.editItems = PermissionScope.All;
                permission.isFolderContact = true;
                permission.isFolderOwner = true;
                permission.isFolderVisible = true;
                permission.readItems = FolderPermissionReadAccess.FullDetails;

                result.Add(FolderPermissionLevel.Owner, permission);

                permission = new FolderPermission();
                permission.canCreateItems = false;
                permission.canCreateSubFolders = false;
                permission.deleteItems = PermissionScope.None;
                permission.editItems = PermissionScope.None;
                permission.isFolderContact = false;
                permission.isFolderOwner = false;
                permission.isFolderVisible = false;
                permission.readItems = FolderPermissionReadAccess.TimeOnly;

                result.Add(FolderPermissionLevel.FreeBusyTimeOnly, permission);

                permission = new FolderPermission();
                permission.canCreateItems = false;
                permission.canCreateSubFolders = false;
                permission.deleteItems = PermissionScope.None;
                permission.editItems = PermissionScope.None;
                permission.isFolderContact = false;
                permission.isFolderOwner = false;
                permission.isFolderVisible = false;
                permission.readItems = FolderPermissionReadAccess.TimeAndSubjectAndLocation;

                result.Add(FolderPermissionLevel.FreeBusyTimeAndSubjectAndLocation, permission);

                return result;
            });

        #endregion

        /// <summary>
        /// Variants of pre-defined permission levels that Outlook also displays with the same levels.
        /// </summary>
        private static LazyMember<List<FolderPermission>> levelVariants = new LazyMember<List<FolderPermission>>(
            delegate()
            {
                List<FolderPermission> results = new List<FolderPermission>();

                FolderPermission permissionNone = FolderPermission.defaultPermissions.Member[FolderPermissionLevel.None];
                FolderPermission permissionOwner = FolderPermission.defaultPermissions.Member[FolderPermissionLevel.Owner];

                // PermissionLevelNoneOption1
                FolderPermission permission = permissionNone.Clone();
                permission.isFolderVisible = true;
                results.Add(permission);

                // PermissionLevelNoneOption2
                permission = permissionNone.Clone();
                permission.isFolderContact = true;
                results.Add(permission);

                // PermissionLevelNoneOption3
                permission = permissionNone.Clone();
                permission.isFolderContact = true;
                permission.isFolderVisible = true;
                results.Add(permission);

                // PermissionLevelOwnerOption1
                permission = permissionOwner.Clone();
                permission.isFolderContact = false;
                results.Add(permission);

                return results;
            });

        private UserId userId;
        private bool canCreateItems;
        private bool canCreateSubFolders;
        private bool isFolderOwner;
        private bool isFolderVisible;
        private bool isFolderContact;
        private PermissionScope editItems;
        private PermissionScope deleteItems;
        private FolderPermissionReadAccess readItems;
        private FolderPermissionLevel permissionLevel;

        /// <summary>
        /// Determines whether the specified folder permission is the same as this one. The comparison
        /// does not take UserId and PermissionLevel into consideration.
        /// </summary>
        /// <param name="permission">The folder permission to compare with this folder permission.</param>
        /// <returns>
        /// True is the specified folder permission is equal to this one, false otherwise.
        /// </returns>
        private bool IsEqualTo(FolderPermission permission)
        {
            return
                this.CanCreateItems == permission.CanCreateItems &&
                this.CanCreateSubFolders == permission.CanCreateSubFolders &&
                this.IsFolderContact == permission.IsFolderContact &&
                this.IsFolderVisible == permission.IsFolderVisible &&
                this.IsFolderOwner == permission.IsFolderOwner &&
                this.EditItems == permission.EditItems &&
                this.DeleteItems == permission.DeleteItems &&
                this.ReadItems == permission.ReadItems;
        }

        /// <summary>
        /// Create a copy of this FolderPermission instance.
        /// </summary>
        /// <returns>
        /// Clone of this instance.
        /// </returns>
        private FolderPermission Clone()
        {
            return (FolderPermission)this.MemberwiseClone();
        }

        /// <summary>
        /// Determines the permission level of this folder permission based on its individual settings,
        /// and sets the PermissionLevel property accordingly.
        /// </summary>
        private void AdjustPermissionLevel()
        {
            foreach (KeyValuePair<FolderPermissionLevel, FolderPermission> keyValuePair in defaultPermissions.Member)
            {
                if (this.IsEqualTo(keyValuePair.Value))
                {
                    this.permissionLevel = keyValuePair.Key;
                    return;
                }
            }

            this.permissionLevel = FolderPermissionLevel.Custom;
        }

        /// <summary>
        /// Copies the values of the individual permissions of the specified folder permission
        /// to this folder permissions.
        /// </summary>
        /// <param name="permission">The folder permission to copy the values from.</param>
        private void AssignIndividualPermissions(FolderPermission permission)
        {
            this.canCreateItems = permission.CanCreateItems;
            this.canCreateSubFolders = permission.CanCreateSubFolders;
            this.isFolderContact = permission.IsFolderContact;
            this.isFolderOwner = permission.IsFolderOwner;
            this.isFolderVisible = permission.IsFolderVisible;
            this.editItems = permission.EditItems;
            this.deleteItems = permission.DeleteItems;
            this.readItems = permission.ReadItems;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FolderPermission"/> class.
        /// </summary>
        public FolderPermission()
            : base()
        {
            this.UserId = new UserId();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FolderPermission"/> class.
        /// </summary>
        /// <param name="userId">The Id of the user  the permission applies to.</param>
        /// <param name="permissionLevel">The level of the permission.</param>
        public FolderPermission(UserId userId, FolderPermissionLevel permissionLevel)
        {
            EwsUtilities.ValidateParam(userId, "userId");

            this.userId = userId;
            this.PermissionLevel = permissionLevel;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FolderPermission"/> class.
        /// </summary>
        /// <param name="primarySmtpAddress">The primary SMTP address of the user the permission applies to.</param>
        /// <param name="permissionLevel">The level of the permission.</param>
        public FolderPermission(string primarySmtpAddress, FolderPermissionLevel permissionLevel)
        {
            this.userId = new UserId(primarySmtpAddress);
            this.PermissionLevel = permissionLevel;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FolderPermission"/> class.
        /// </summary>
        /// <param name="standardUser">The standard user the permission applies to.</param>
        /// <param name="permissionLevel">The level of the permission.</param>
        public FolderPermission(StandardUser standardUser, FolderPermissionLevel permissionLevel)
        {
            this.userId = new UserId(standardUser);
            this.PermissionLevel = permissionLevel;
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        /// <param name="isCalendarFolder">if set to <c>true</c> calendar permissions are allowed.</param>
        /// <param name="permissionIndex">Index of the permission.</param>
        internal void Validate(bool isCalendarFolder, int permissionIndex)
        {
            // Check UserId
            if (!this.UserId.IsValid())
            {
                throw new ServiceValidationException(
                    string.Format(
                        Strings.FolderPermissionHasInvalidUserId,
                        permissionIndex));
            }

            // If this permission is to be used for a non-calendar folder make sure that read access and permission level aren't set to Calendar-only values
            if (!isCalendarFolder)
            {
                if ((this.readItems == FolderPermissionReadAccess.TimeAndSubjectAndLocation) ||
                    (this.readItems == FolderPermissionReadAccess.TimeOnly))
                {
                    throw new ServiceLocalException(
                        string.Format(
                            Strings.ReadAccessInvalidForNonCalendarFolder,
                            this.readItems));
                }

                if ((this.permissionLevel == FolderPermissionLevel.FreeBusyTimeAndSubjectAndLocation) ||
                    (this.permissionLevel == FolderPermissionLevel.FreeBusyTimeOnly))
                {
                    throw new ServiceLocalException(
                        string.Format(
                            Strings.PermissionLevelInvalidForNonCalendarFolder,
                            this.permissionLevel));
                }
            }
        }

        /// <summary>
        /// Gets the Id of the user the permission applies to.
        /// </summary>
        public UserId UserId
        {
            get
            { 
                return this.userId;
            }

            set
            {
                if (this.userId != null)
                {
                    this.userId.OnChange -= this.PropertyChanged;
                }

                this.SetFieldValue<UserId>(ref this.userId, value);

                if (this.userId != null)
                {
                    this.userId.OnChange += this.PropertyChanged;
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the user can create new items.
        /// </summary>
        public bool CanCreateItems
        {
            get
            { 
                return this.canCreateItems;
            }

            set
            { 
                this.SetFieldValue<bool>(ref this.canCreateItems, value);
                this.AdjustPermissionLevel();
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the user can create sub-folders.
        /// </summary>
        public bool CanCreateSubFolders
        {
            get
            { 
                return this.canCreateSubFolders;
            }

            set
            { 
                this.SetFieldValue<bool>(ref this.canCreateSubFolders, value);
                this.AdjustPermissionLevel();
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the user owns the folder.
        /// </summary>
        public bool IsFolderOwner
        {
            get
            {
                return this.isFolderOwner;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.isFolderOwner, value);
                this.AdjustPermissionLevel();
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the folder is visible to the user.
        /// </summary>
        public bool IsFolderVisible
        {
            get
            {
                return this.isFolderVisible;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.isFolderVisible, value);
                this.AdjustPermissionLevel();
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the user is a contact for the folder.
        /// </summary>
        public bool IsFolderContact
        {
            get
            {
                return this.isFolderContact;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.isFolderContact, value);
                this.AdjustPermissionLevel();
            }
        }

        /// <summary>
        /// Gets or sets a value indicating if/how the user can edit existing items.
        /// </summary>
        public PermissionScope EditItems
        {
            get
            {
                return this.editItems;
            }

            set
            {
                this.SetFieldValue<PermissionScope>(ref this.editItems, value);
                this.AdjustPermissionLevel();
            }
        }

        /// <summary>
        /// Gets or sets a value indicating if/how the user can delete existing items.
        /// </summary>
        public PermissionScope DeleteItems
        {
            get
            {
                return this.deleteItems;
            }

            set
            {
                this.SetFieldValue<PermissionScope>(ref this.deleteItems, value);
                this.AdjustPermissionLevel();
            }
        }

        /// <summary>
        /// Gets or sets the read items access permission.
        /// </summary>
        public FolderPermissionReadAccess ReadItems
        {
            get
            {
                return this.readItems;
            }

            set
            {
                this.SetFieldValue<FolderPermissionReadAccess>(ref this.readItems, value);
                this.AdjustPermissionLevel();
            }
        }

        /// <summary>
        /// Gets or sets the permission level.
        /// </summary>
        public FolderPermissionLevel PermissionLevel
        {
            get
            { 
                return this.permissionLevel;
            }

            set
            {
                if (this.permissionLevel != value)
                {
                    if (value == FolderPermissionLevel.Custom)
                    {
                        throw new ServiceLocalException(Strings.CannotSetPermissionLevelToCustom);
                    }

                    this.AssignIndividualPermissions(defaultPermissions.Member[value]);
                    this.SetFieldValue<FolderPermissionLevel>(ref this.permissionLevel, value);
                }
            }
        }

        /// <summary>
        /// Gets the permission level that Outlook would display for this folder permission.
        /// </summary>
        public FolderPermissionLevel DisplayPermissionLevel
        {
            get
            {
                // If permission level is set to Custom, see if there's a variant
                // that Outlook would map to the same permission level.
                if (this.permissionLevel == FolderPermissionLevel.Custom)
                {
                    foreach (FolderPermission variant in FolderPermission.levelVariants.Member)
                    {
                        if (this.IsEqualTo(variant))
                        {
                            return variant.PermissionLevel;
                        }
                    }
                }

                return this.permissionLevel;
            }
        }

        /// <summary>
        /// Property was changed.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        private void PropertyChanged(ComplexProperty complexProperty)
        {
            this.Changed();
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
                case XmlElementNames.UserId:
                    this.UserId = new UserId();
                    this.UserId.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.CanCreateItems:
                    this.canCreateItems = reader.ReadValue<bool>();
                    return true;
                case XmlElementNames.CanCreateSubFolders:
                    this.canCreateSubFolders = reader.ReadValue<bool>();
                    return true;
                case XmlElementNames.IsFolderOwner:
                    this.isFolderOwner = reader.ReadValue<bool>();
                    return true;
                case XmlElementNames.IsFolderVisible:
                    this.isFolderVisible = reader.ReadValue<bool>();
                    return true;
                case XmlElementNames.IsFolderContact:
                    this.isFolderContact = reader.ReadValue<bool>();
                    return true;
                case XmlElementNames.EditItems:
                    this.editItems = reader.ReadValue<PermissionScope>();
                    return true;
                case XmlElementNames.DeleteItems:
                    this.deleteItems = reader.ReadValue<PermissionScope>();
                    return true;
                case XmlElementNames.ReadItems:
                    this.readItems = reader.ReadValue<FolderPermissionReadAccess>();
                    return true;
                case XmlElementNames.PermissionLevel:
                case XmlElementNames.CalendarPermissionLevel:
                    this.permissionLevel = reader.ReadValue<FolderPermissionLevel>();
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal override void LoadFromXml(
            EwsServiceXmlReader reader,
            XmlNamespace xmlNamespace,
            string xmlElementName)
        {
            base.LoadFromXml(reader, xmlNamespace, xmlElementName);

            this.AdjustPermissionLevel();
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="isCalendarFolder">If true, this permission is for a calendar folder.</param>
        internal void WriteElementsToXml(EwsServiceXmlWriter writer, bool isCalendarFolder)
        {
            if (this.UserId != null)
            {
                this.UserId.WriteToXml(writer, XmlElementNames.UserId);
            }

            if (this.PermissionLevel == FolderPermissionLevel.Custom)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.CanCreateItems,
                    this.CanCreateItems);

                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.CanCreateSubFolders,
                    this.CanCreateSubFolders);

                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.IsFolderOwner,
                    this.IsFolderOwner);

                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.IsFolderVisible,
                    this.IsFolderVisible);

                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.IsFolderContact,
                    this.IsFolderContact);

                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.EditItems,
                    this.EditItems);

                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.DeleteItems,
                    this.DeleteItems);

                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.ReadItems,
                    this.ReadItems);
            }

            writer.WriteElementValue(
                XmlNamespace.Types,
                isCalendarFolder ? XmlElementNames.CalendarPermissionLevel : XmlElementNames.PermissionLevel,
                this.PermissionLevel);
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="isCalendarFolder">If true, this permission is for a calendar folder.</param>
        internal void WriteToXml(
            EwsServiceXmlWriter writer,
            string xmlElementName,
            bool isCalendarFolder)
        {
            writer.WriteStartElement(this.Namespace, xmlElementName);
            this.WriteAttributesToXml(writer);
            this.WriteElementsToXml(writer, isCalendarFolder);
            writer.WriteEndElement();
        }
    }
}