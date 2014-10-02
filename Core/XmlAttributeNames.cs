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
// <summary>Defines the XmlAttributeNames class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// XML attribute names.
    /// </summary>
    internal static class XmlAttributeNames
    {
        public const string XmlNs = "xmlns";
        public const string Id = "Id";
        public const string ChangeKey = "ChangeKey";
        public const string RecurringMasterId = "RecurringMasterId";
        public const string InstanceIndex = "InstanceIndex";
        public const string OccurrenceId = "OccurrenceId";
        public const string Traversal = "Traversal";
        public const string ViewFilter = "ViewFilter";
        public const string Offset = "Offset";
        public const string MaxEntriesReturned = "MaxEntriesReturned";
        public const string BasePoint = "BasePoint";
        public const string ResponseClass = "ResponseClass";
        public const string IndexedPagingOffset = "IndexedPagingOffset";
        public const string TotalItemsInView = "TotalItemsInView";
        public const string IncludesLastItemInRange = "IncludesLastItemInRange";
        public const string BodyType = "BodyType";
        public const string MessageDisposition = "MessageDisposition";
        public const string SaveItemToFolder = "SaveItemToFolder";
        public const string RootItemChangeKey = "RootItemChangeKey";
        public const string DeleteType = "DeleteType";
        public const string DeleteSubFolders = "DeleteSubFolders";
        public const string AffectedTaskOccurrences = "AffectedTaskOccurrences";
        public const string SendMeetingCancellations = "SendMeetingCancellations";
        public const string SuppressReadReceipts = XmlElementNames.SuppressReadReceipts;
        public const string FieldURI = "FieldURI";
        public const string FieldIndex = "FieldIndex";
        public const string ConflictResolution = "ConflictResolution";
        public const string SendMeetingInvitationsOrCancellations = "SendMeetingInvitationsOrCancellations";
        public const string CharacterSet = "CharacterSet";
        public const string HeaderName = "HeaderName";
        public const string SendMeetingInvitations = "SendMeetingInvitations";
        public const string Key = "Key";
        public const string RoutingType = "RoutingType";
        public const string MailboxType = "MailboxType";
        public const string DistinguishedPropertySetId = "DistinguishedPropertySetId";
        public const string PropertySetId = "PropertySetId";
        public const string PropertyTag = "PropertyTag";
        public const string PropertyName = "PropertyName";
        public const string PropertyId = "PropertyId";
        public const string PropertyType = "PropertyType";
        public const string TimeZoneName = "TimeZoneName";
        public const string ReturnFullContactData = "ReturnFullContactData";
        public const string ContactDataShape = "ContactDataShape";
        public const string Numerator = "Numerator";
        public const string Denominator = "Numerator";
        public const string Value = "Value";
        public const string ContainmentMode = "ContainmentMode";
        public const string ContainmentComparison = "ContainmentComparison";
        public const string Order = "Order";
        public const string StartDate = "StartDate";
        public const string EndDate = "EndDate";
        public const string Version = "Version";
        public const string Aggregate = "Aggregate";
        public const string SearchScope = "SearchScope";
        public const string Format = "Format";
        public const string Mailbox = "Mailbox";
        public const string DestinationFormat = "DestinationFormat";
        public const string FolderId = "FolderId";
        public const string ItemId = "ItemId";
        public const string IncludePermissions = "IncludePermissions";
        public const string InitialName = "InitialName";
        public const string FinalName = "FinalName";
        public const string AuthenticationMethod = "AuthenticationMethod";
        public const string Time = "Time";
        public const string Name = "Name";
        public const string Bias = "Bias";
        public const string Kind = "Kind";
        public const string SubscribeToAllFolders = "SubscribeToAllFolders";
        public const string PublicFolderServer = "PublicFolderServer";
        public const string IsArchive = "IsArchive";
        public const string ReturnHighlightTerms = "ReturnHighlightTerms";
        public const string IsExplicit = "IsExplicit";
        public const string ClientExtensionUserIdentity = "UserId";
        public const string ClientExtensionEnabledOnly = "EnabledOnly";
        public const string SetClientExtensionActionId = "ActionId";
        public const string ClientExtensionId = "ExtensionId";
        public const string ClientExtensionIsAvailable = "IsAvailable";
        public const string ClientExtensionIsMandatory = "IsMandatory";
        public const string ClientExtensionIsEnabledByDefault = "IsEnabledByDefault";
        public const string ClientExtensionProvidedTo = "ProvidedTo";
        public const string ClientExtensionType = "Type";
        public const string ClientExtensionScope = "Scope";
        public const string ClientExtensionMarketplaceAssetID = "MarketplaceAssetId";
        public const string ClientExtensionMarketplaceContentMarket = "MarketplaceContentMarket";
        public const string ClientExtensionAppStatus = "AppStatus";
        public const string ClientExtensionEtoken = "Etoken";
        public const string IsTruncated = "IsTruncated";
        public const string IsJunk = "IsJunk";
        public const string MoveItem = "MoveItem";

        // xsi attributes
        public const string Nil = "nil";
        public const string Type = "type";
    }
}
