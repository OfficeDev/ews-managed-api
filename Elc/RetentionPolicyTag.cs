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
// <summary>Defines the RetentionPolicyTag class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// Represents retention policy tag object.
    /// </summary>
    public sealed class RetentionPolicyTag
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public RetentionPolicyTag()
        {
        }

        /// <summary>
        /// Constructor for retention policy tag.
        /// </summary>
        /// <param name="displayName">Display name.</param>
        /// <param name="retentionId">Retention id.</param>
        /// <param name="retentionPeriod">Retention period.</param>
        /// <param name="type">Retention folder type.</param>
        /// <param name="retentionAction">Retention action.</param>
        /// <param name="isVisible">Is visible.</param>
        /// <param name="optedInto">Opted into.</param>
        /// <param name="isArchive">Is archive tag.</param>
        internal RetentionPolicyTag(
            string displayName,
            Guid retentionId,
            int retentionPeriod,
            ElcFolderType type,
            RetentionActionType retentionAction,
            bool isVisible,
            bool optedInto,
            bool isArchive)
        {
            DisplayName = displayName;
            RetentionId = retentionId;
            RetentionPeriod = retentionPeriod;
            Type = type;
            RetentionAction = retentionAction;
            IsVisible = isVisible;
            OptedInto = optedInto;
            IsArchive = isArchive;
        }

        /// <summary>
        /// Load from xml.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Retention policy tag object.</returns>
        internal static RetentionPolicyTag LoadFromXml(EwsServiceXmlReader reader)
        {
            reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, XmlElementNames.RetentionPolicyTag);

            RetentionPolicyTag retentionPolicyTag = new RetentionPolicyTag();
            retentionPolicyTag.DisplayName = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.DisplayName);
            retentionPolicyTag.RetentionId = new Guid(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.RetentionId));
            retentionPolicyTag.RetentionPeriod = reader.ReadElementValue<int>(XmlNamespace.Types, XmlElementNames.RetentionPeriod);
            retentionPolicyTag.Type = reader.ReadElementValue<ElcFolderType>(XmlNamespace.Types, XmlElementNames.Type);
            retentionPolicyTag.RetentionAction = reader.ReadElementValue<RetentionActionType>(XmlNamespace.Types, XmlElementNames.RetentionAction);

            // Description is not a required property.
            reader.Read();
            if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Description))
            {
                retentionPolicyTag.Description = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Description);
            }

            retentionPolicyTag.IsVisible = reader.ReadElementValue<bool>(XmlNamespace.Types, XmlElementNames.IsVisible);
            retentionPolicyTag.OptedInto = reader.ReadElementValue<bool>(XmlNamespace.Types, XmlElementNames.OptedInto);
            retentionPolicyTag.IsArchive = reader.ReadElementValue<bool>(XmlNamespace.Types, XmlElementNames.IsArchive);

            return retentionPolicyTag;
        }

        /// <summary>
        /// Load from json.
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        /// <returns>Retention policy tag object.</returns>
        internal static RetentionPolicyTag LoadFromJson(JsonObject jsonObject)
        {
            RetentionPolicyTag retentionPolicyTag = new RetentionPolicyTag();

            if (jsonObject.ContainsKey(XmlElementNames.DisplayName))
            {
                retentionPolicyTag.DisplayName = jsonObject.ReadAsString(XmlElementNames.DisplayName);
            }

            if (jsonObject.ContainsKey(XmlElementNames.RetentionId))
            {
                retentionPolicyTag.RetentionId = new Guid(jsonObject.ReadAsString(XmlElementNames.RetentionId));
            }

            if (jsonObject.ContainsKey(XmlElementNames.RetentionPeriod))
            {
                retentionPolicyTag.RetentionPeriod = jsonObject.ReadAsInt(XmlElementNames.RetentionPeriod);
            }

            if (jsonObject.ContainsKey(XmlElementNames.Type))
            {
                retentionPolicyTag.Type = jsonObject.ReadEnumValue<ElcFolderType>(XmlElementNames.Type);
            }

            if (jsonObject.ContainsKey(XmlElementNames.RetentionAction))
            {
                retentionPolicyTag.RetentionAction = jsonObject.ReadEnumValue<RetentionActionType>(XmlElementNames.RetentionAction);
            }

            if (jsonObject.ContainsKey(XmlElementNames.Description))
            {
                retentionPolicyTag.Description = jsonObject.ReadAsString(XmlElementNames.Description);
            }

            if (jsonObject.ContainsKey(XmlElementNames.IsVisible))
            {
                retentionPolicyTag.IsVisible = jsonObject.ReadAsBool(XmlElementNames.IsVisible);
            }

            if (jsonObject.ContainsKey(XmlElementNames.OptedInto))
            {
                retentionPolicyTag.OptedInto = jsonObject.ReadAsBool(XmlElementNames.OptedInto);
            }

            if (jsonObject.ContainsKey(XmlElementNames.IsArchive))
            {
                retentionPolicyTag.IsArchive = jsonObject.ReadAsBool(XmlElementNames.IsArchive);
            }

            return retentionPolicyTag;
        }

        /// <summary>
        /// Retention policy tag display name.
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// Retention Id.
        /// </summary>
        public Guid RetentionId { get; set; }

        /// <summary>
        /// Retention period in time span.
        /// </summary>
        public int RetentionPeriod { get; set; }

        /// <summary>
        /// Retention type.
        /// </summary>
        public ElcFolderType Type { get; set; }

        /// <summary>
        /// Retention action.
        /// </summary>
        public RetentionActionType RetentionAction { get; set; }

        /// <summary>
        /// Retention policy tag description.
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// Is this a visible tag?
        /// </summary>
        public bool IsVisible { get; set; }

        /// <summary>
        /// Is this a opted into tag?
        /// </summary>
        public bool OptedInto { get; set; }

        /// <summary>
        /// Is this an archive tag?
        /// </summary>
        public bool IsArchive { get; set; }
    }
}
