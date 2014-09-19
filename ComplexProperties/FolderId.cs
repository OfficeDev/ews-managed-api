// ---------------------------------------------------------------------------
// <copyright file="FolderId.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the FolderId class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents the Id of a folder.
    /// </summary>
    public sealed class FolderId : ServiceId
    {
        private WellKnownFolderName? folderName;
        private Mailbox mailbox;

        /// <summary>
        /// Initializes a new instance of the <see cref="FolderId"/> class.
        /// </summary>
        internal FolderId()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FolderId"/> class. Use this constructor
        /// to link this FolderId to an existing folder that you have the unique Id of.
        /// </summary>
        /// <param name="uniqueId">The unique Id used to initialize the FolderId.</param>
        public FolderId(string uniqueId)
            : base(uniqueId)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FolderId"/> class. Use this constructor
        /// to link this FolderId to a well known folder (e.g. Inbox, Calendar or Contacts).
        /// </summary>
        /// <param name="folderName">The folder name used to initialize the FolderId.</param>
        public FolderId(WellKnownFolderName folderName)
            : base()
        {
            this.folderName = folderName;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FolderId"/> class. Use this constructor
        /// to link this FolderId to a well known folder (e.g. Inbox, Calendar or Contacts) in a
        /// specific mailbox.
        /// </summary>
        /// <param name="folderName">The folder name used to initialize the FolderId.</param>
        /// <param name="mailbox">The mailbox used to initialize the FolderId.</param>
        public FolderId(WellKnownFolderName folderName, Mailbox mailbox)
            : this(folderName)
        {
            this.mailbox = mailbox;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return this.FolderName.HasValue ? XmlElementNames.DistinguishedFolderId : XmlElementNames.FolderId;
        }

        /// <summary>
        /// Writes attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            if (this.FolderName.HasValue)
            {
                writer.WriteAttributeValue(XmlAttributeNames.Id, this.FolderName.Value.ToString().ToLowerInvariant());

                if (this.Mailbox != null)
                {
                    this.Mailbox.WriteToXml(writer, XmlElementNames.Mailbox);
                }
            }
            else
            {
                base.WriteAttributesToXml(writer);
            }
        }

        /// <summary>
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service"></param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object InternalToJson(ExchangeService service)
        {
            if (this.FolderName.HasValue)
            {
                JsonObject jsonProperty = new JsonObject();

                jsonProperty.AddTypeParameter(this.GetXmlElementName());
                jsonProperty.Add(XmlAttributeNames.Id, this.FolderName.Value.ToString().ToLowerInvariant());

                if (this.Mailbox != null)
                {
                    jsonProperty.Add(XmlElementNames.Mailbox, this.Mailbox.InternalToJson(service));
                }

                return jsonProperty;
            }
            else
            {
                return base.InternalToJson(service);
            }
        }

        /// <summary>
        /// Validates FolderId against a specified request version.
        /// </summary>
        /// <param name="version">The version.</param>
        internal void Validate(ExchangeVersion version)
        {
            // The FolderName property is a WellKnownFolderName, an enumeration type. If the property
            // is set, make sure that the value is valid for the request version.
            if (this.FolderName.HasValue)
            {
                EwsUtilities.ValidateEnumVersionValue(this.FolderName.Value, version);
            }
        }

        /// <summary>
        /// Gets the name of the folder associated with the folder Id. Name and Id are mutually exclusive; if one is set, the other is null.
        /// </summary>
        public WellKnownFolderName? FolderName
        {
            get { return this.folderName; }
        }

        /// <summary>
        /// Gets the mailbox of the folder. Mailbox is only set when FolderName is set.
        /// </summary>
        public Mailbox Mailbox
        {
            get { return this.mailbox; }
        }

        /// <summary>
        ///  Defines an implicit conversion between string and FolderId.
        /// </summary>
        /// <param name="uniqueId">The unique Id to convert to FolderId.</param>
        /// <returns>A FolderId initialized with the specified unique Id.</returns>
        public static implicit operator FolderId(string uniqueId)
        {
            return new FolderId(uniqueId);
        }

        /// <summary>
        /// Defines an implicit conversion between WellKnownFolderName and FolderId.
        /// </summary>
        /// <param name="folderName">The folder name to convert to FolderId.</param>
        /// <returns>A FolderId initialized with the specified folder name.</returns>
        public static implicit operator FolderId(WellKnownFolderName folderName)
        {
            return new FolderId(folderName);
        }

        /// <summary>
        /// True if this instance is valid, false otherthise.
        /// </summary>
        /// <value><c>true</c> if this instance is valid; otherwise, <c>false</c>.</value>
        internal override bool IsValid
        {
            get
            {
                if (this.FolderName.HasValue)
                {
                    return (this.Mailbox == null) || this.Mailbox.IsValid;
                }
                else
                {
                    return base.IsValid;
                }
            }
        }

        /// <summary>
        /// Determines whether the specified <see cref="T:System.Object"/> is equal to the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <param name="obj">The <see cref="T:System.Object"/> to compare with the current <see cref="T:System.Object"/>.</param>
        /// <returns>
        /// true if the specified <see cref="T:System.Object"/> is equal to the current <see cref="T:System.Object"/>; otherwise, false.
        /// </returns>
        /// <exception cref="T:System.NullReferenceException">The <paramref name="obj"/> parameter is null.</exception>
        public override bool Equals(object obj)
        {
            if (object.ReferenceEquals(this, obj))
            {
                return true;
            }
            else
            {
                FolderId other = obj as FolderId;

                if (other == null)
                {
                    return false;
                }
                else if (this.FolderName.HasValue)
                {
                    if (other.FolderName.HasValue && this.FolderName.Value.Equals(other.FolderName.Value))
                    {
                        if (this.Mailbox != null)
                        {
                            return this.Mailbox.Equals(other.Mailbox);
                        }
                        else if (other.Mailbox == null)
                        {
                            return true;
                        }
                    }
                }
                else if (base.Equals(other))
                {
                    return true;
                }

                return false;
            }
        }

        /// <summary>
        /// Serves as a hash function for a particular type.
        /// </summary>
        /// <returns>
        /// A hash code for the current <see cref="T:System.Object"/>.
        /// </returns>
        public override int GetHashCode()
        {
            int hashCode;

            if (this.FolderName.HasValue)
            {
                hashCode = this.FolderName.Value.GetHashCode();

                if ((this.Mailbox != null) && this.Mailbox.IsValid)
                {
                    hashCode = hashCode ^ this.Mailbox.GetHashCode();
                }
            }
            else
            {
                hashCode = base.GetHashCode();
            }

            return hashCode;
        }

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override string ToString()
        {
            if (this.IsValid)
            {
                if (this.FolderName.HasValue)
                {
                    if ((this.Mailbox != null) && mailbox.IsValid)
                    {
                        return string.Format("{0} ({1})", this.folderName.Value, this.Mailbox.ToString());
                    }
                    else
                    {
                        return this.FolderName.Value.ToString();
                    }
                }
                else
                {
                    return base.ToString();
                }
            }
            else
            {
                return string.Empty;
            }
        }
    }
}
