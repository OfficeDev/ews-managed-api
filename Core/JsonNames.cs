// ---------------------------------------------------------------------------
// <copyright file="JsonNames.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the XmlElementNames class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// JSON names not shared with the XmlElementNames or XmlAttributeNames classes.
    /// </summary>
    internal static class JsonNames
    {
        public const string Events = "Events";
        public const string NotificationType = "NotificationType";
        public const string OldFolderId = "OldFolderId";
        public const string OldItemId = "OldItemId";
        public const string PathToExtendedFieldType = "ExtendedPropertyUri";
        public const string PathToIndexedFieldType = "DictionaryPropertyUri";
        public const string PathToUnindexedFieldType = "PropertyUri";
        public const string Path = "Path";
        public const string RecurrencePattern = "RecurrencePattern";
        public const string RecurrenceRange = "RecurrenceRange";
    }
}
