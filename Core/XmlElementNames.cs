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
    /// XML element names.
    /// </summary>
    internal static class XmlElementNames
    {
        public const string AllProperties = "AllProperties";
        public const string ParentFolderIds = "ParentFolderIds";
        public const string DistinguishedFolderId = "DistinguishedFolderId";
        public const string ItemId = "ItemId";
        public const string ItemIds = "ItemIds";
        public const string FolderId = "FolderId";
        public const string FolderIds = "FolderIds";
        public const string SourceId = "SourceId";
        public const string OccurrenceItemId = "OccurrenceItemId";
        public const string RecurringMasterItemId = "RecurringMasterItemId";
        public const string ItemShape = "ItemShape";
        public const string FolderShape = "FolderShape";
        public const string BaseShape = "BaseShape";
        public const string IndexedPageItemView = "IndexedPageItemView";
        public const string IndexedPageFolderView = "IndexedPageFolderView";
        public const string FractionalPageItemView = "FractionalPageItemView";
        public const string FractionalPageFolderView = "FractionalPageFolderView";
        public const string SeekToConditionPageItemView = "SeekToConditionPageItemView";
        public const string ResponseCode = "ResponseCode";
        public const string RootFolder = "RootFolder";
        public const string Folder = "Folder";
        public const string ContactsFolder = "ContactsFolder";
        public const string TasksFolder = "TasksFolder";
        public const string SearchFolder = "SearchFolder";
        public const string Folders = "Folders";
        public const string Item = "Item";
        public const string Items = "Items";
        public const string Message = "Message";
        public const string Mailbox = "Mailbox";
        public const string Body = "Body";
        public const string From = "From";
        public const string Sender = "Sender";
        public const string Name = "Name";
        public const string Address = "Address";
        public const string EmailAddress = "EmailAddress";
        public const string RoutingType = "RoutingType";
        public const string MailboxType = "MailboxType";
        public const string ToRecipients = "ToRecipients";
        public const string CcRecipients = "CcRecipients";
        public const string BccRecipients = "BccRecipients";
        public const string ReplyTo = "ReplyTo";
        public const string ConversationTopic = "ConversationTopic";
        public const string ConversationIndex = "ConversationIndex";
        public const string IsDeliveryReceiptRequested = "IsDeliveryReceiptRequested";
        public const string IsRead = "IsRead";
        public const string IsReadReceiptRequested = "IsReadReceiptRequested";
        public const string IsResponseRequested = "IsResponseRequested";
        public const string InternetMessageId = "InternetMessageId";
        public const string References = "References";
        public const string ParentItemId = "ParentItemId";
        public const string ParentFolderId = "ParentFolderId";
        public const string ChildFolderCount = "ChildFolderCount";
        public const string DisplayName = "DisplayName";
        public const string TotalCount = "TotalCount";
        public const string ItemClass = "ItemClass";
        public const string FolderClass = "FolderClass";
        public const string Subject = "Subject";
        public const string MimeContent = "MimeContent";
        public const string MimeContentUTF8 = "MimeContentUTF8";
        public const string Sensitivity = "Sensitivity";
        public const string Attachments = "Attachments";
        public const string DateTimeReceived = "DateTimeReceived";
        public const string Size = "Size";
        public const string Categories = "Categories";
        public const string Importance = "Importance";
        public const string InReplyTo = "InReplyTo";
        public const string IsSubmitted = "IsSubmitted";
        public const string IsAssociated = "IsAssociated";
        public const string IsDraft = "IsDraft";
        public const string IsFromMe = "IsFromMe";
        public const string IsHidden = "IsHidden";
        public const string IsQuickContact = "IsQuickContact";
        public const string IsResend = "IsResend";
        public const string IsUnmodified = "IsUnmodified";
        public const string IsWritable = "IsWritable";
        public const string InternetMessageHeader = "InternetMessageHeader";
        public const string InternetMessageHeaders = "InternetMessageHeaders";
        public const string DateTimeSent = "DateTimeSent";
        public const string DateTimeCreated = "DateTimeCreated";
        public const string ResponseObjects = "ResponseObjects";
        public const string ReminderDueBy = "ReminderDueBy";
        public const string ReminderIsSet = "ReminderIsSet";
        public const string ReminderMinutesBeforeStart = "ReminderMinutesBeforeStart";
        public const string DisplayCc = "DisplayCc";
        public const string DisplayTo = "DisplayTo";
        public const string HasAttachments = "HasAttachments";
        public const string ExtendedProperty = "ExtendedProperty";
        public const string Culture = "Culture";
        public const string FileAttachment = "FileAttachment";
        public const string ItemAttachment = "ItemAttachment";
        public const string ReferenceAttachment = "ReferenceAttachment";
        public const string AttachLongPathName = "AttachLongPathName";
        public const string ProviderType = "ProviderType";
        public const string ProviderEndpointUrl = "ProviderEndpointUrl";
        public const string AttachmentThumbnailUrl = "AttachmentThumbnailUrl";
        public const string AttachmentPreviewUrl = "AttachmentPreviewUrl";
        public const string PermissionType = "PermissionType";
        public const string AttachmentIsFolder = "AttachmentIsFolder";
        public const string AttachmentIds = "AttachmentIds";
        public const string AttachmentId = "AttachmentId";
        public const string ContentType = "ContentType";
        public const string ContentLocation = "ContentLocation";
        public const string ContentId = "ContentId";
        public const string Content = "Content";
        public const string SavedItemFolderId = "SavedItemFolderId";
        public const string MessageText = "MessageText";
        public const string DescriptiveLinkKey = "DescriptiveLinkKey";
        public const string ItemChange = "ItemChange";
        public const string ItemChanges = "ItemChanges";
        public const string FolderChange = "FolderChange";
        public const string FolderChanges = "FolderChanges";
        public const string Updates = "Updates";
        public const string AppendToItemField = "AppendToItemField";
        public const string SetItemField = "SetItemField";
        public const string DeleteItemField = "DeleteItemField";
        public const string SetFolderField = "SetFolderField";
        public const string DeleteFolderField = "DeleteFolderField";
        public const string FieldURI = "FieldURI";
        public const string RootItemId = "RootItemId";
        public const string ReferenceItemId = "ReferenceItemId";
        public const string NewBodyContent = "NewBodyContent";
        public const string ReplyToItem = "ReplyToItem";
        public const string ReplyAllToItem = "ReplyAllToItem";
        public const string ForwardItem = "ForwardItem";
        public const string AcceptItem = "AcceptItem";
        public const string TentativelyAcceptItem = "TentativelyAcceptItem";
        public const string DeclineItem = "DeclineItem";
        public const string CancelCalendarItem = "CancelCalendarItem";
        public const string RemoveItem = "RemoveItem";
        public const string SuppressReadReceipt = "SuppressReadReceipt";
        public const string SuppressReadReceipts = "SuppressReadReceipts";
        public const string String = "String";
        public const string Start = "Start";
        public const string End = "End";
        public const string ProposedStart = "ProposedStart";
        public const string ProposedEnd = "ProposedEnd";
        public const string OriginalStart = "OriginalStart";
        public const string IsAllDayEvent = "IsAllDayEvent";
        public const string LegacyFreeBusyStatus = "LegacyFreeBusyStatus";
        public const string Location = "Location";
        public const string When = "When";
        public const string IsMeeting = "IsMeeting";
        public const string IsCancelled = "IsCancelled";
        public const string IsRecurring = "IsRecurring";
        public const string MeetingRequestWasSent = "MeetingRequestWasSent";
        public const string CalendarItemType = "CalendarItemType";
        public const string MyResponseType = "MyResponseType";
        public const string Organizer = "Organizer";
        public const string RequiredAttendees = "RequiredAttendees";
        public const string OptionalAttendees = "OptionalAttendees";
        public const string Resources = "Resources";
        public const string ConflictingMeetingCount = "ConflictingMeetingCount";
        public const string AdjacentMeetingCount = "AdjacentMeetingCount";
        public const string ConflictingMeetings = "ConflictingMeetings";
        public const string AdjacentMeetings = "AdjacentMeetings";
        public const string Duration = "Duration";
        public const string TimeZone = "TimeZone";
        public const string AppointmentReplyTime = "AppointmentReplyTime";
        public const string AppointmentSequenceNumber = "AppointmentSequenceNumber";
        public const string AppointmentState = "AppointmentState";
        public const string Recurrence = "Recurrence";
        public const string FirstOccurrence = "FirstOccurrence";
        public const string LastOccurrence = "LastOccurrence";
        public const string ModifiedOccurrences = "ModifiedOccurrences";
        public const string DeletedOccurrences = "DeletedOccurrences";
        public const string MeetingTimeZone = "MeetingTimeZone";
        public const string ConferenceType = "ConferenceType";
        public const string AllowNewTimeProposal = "AllowNewTimeProposal";
        public const string IsOnlineMeeting = "IsOnlineMeeting";
        public const string MeetingWorkspaceUrl = "MeetingWorkspaceUrl";
        public const string NetShowUrl = "NetShowUrl";
        public const string JoinOnlineMeetingUrl = "JoinOnlineMeetingUrl";
        public const string OnlineMeetingSettings = "OnlineMeetingSettings";
        public const string LobbyBypass = "LobbyBypass";
        public const string AccessLevel = "AccessLevel";
        public const string Presenters = "Presenters";
        public const string CalendarItem = "CalendarItem";
        public const string CalendarFolder = "CalendarFolder";
        public const string Attendee = "Attendee";
        public const string ResponseType = "ResponseType";
        public const string LastResponseTime = "LastResponseTime";
        public const string Occurrence = "Occurrence";
        public const string DeletedOccurrence = "DeletedOccurrence";
        public const string RelativeYearlyRecurrence = "RelativeYearlyRecurrence";
        public const string AbsoluteYearlyRecurrence = "AbsoluteYearlyRecurrence";
        public const string RelativeMonthlyRecurrence = "RelativeMonthlyRecurrence";
        public const string AbsoluteMonthlyRecurrence = "AbsoluteMonthlyRecurrence";
        public const string WeeklyRecurrence = "WeeklyRecurrence";
        public const string DailyRecurrence = "DailyRecurrence";
        public const string DailyRegeneration = "DailyRegeneration";
        public const string WeeklyRegeneration = "WeeklyRegeneration";
        public const string MonthlyRegeneration = "MonthlyRegeneration";
        public const string YearlyRegeneration = "YearlyRegeneration";
        public const string NoEndRecurrence = "NoEndRecurrence";
        public const string EndDateRecurrence = "EndDateRecurrence";
        public const string NumberedRecurrence = "NumberedRecurrence";
        public const string Interval = "Interval";
        public const string DayOfMonth = "DayOfMonth";
        public const string DayOfWeek = "DayOfWeek";
        public const string DaysOfWeek = "DaysOfWeek";
        public const string DayOfWeekIndex = "DayOfWeekIndex";
        public const string Month = "Month";
        public const string StartDate = "StartDate";
        public const string EndDate = "EndDate";
        public const string StartTime = "StartTime";
        public const string EndTime = "EndTime";
        public const string NumberOfOccurrences = "NumberOfOccurrences";
        public const string AssociatedCalendarItemId = "AssociatedCalendarItemId";
        public const string IsDelegated = "IsDelegated";
        public const string IsOutOfDate = "IsOutOfDate";
        public const string HasBeenProcessed = "HasBeenProcessed";
        public const string IsOrganizer = "IsOrganizer";
        public const string MeetingMessage = "MeetingMessage";
        public const string FileAs = "FileAs";
        public const string FileAsMapping = "FileAsMapping";
        public const string GivenName = "GivenName";
        public const string Initials = "Initials";
        public const string MiddleName = "MiddleName";
        public const string NickName = "Nickname";
        public const string CompleteName = "CompleteName";
        public const string CompanyName = "CompanyName";
        public const string EmailAddresses = "EmailAddresses";
        public const string PhysicalAddresses = "PhysicalAddresses";
        public const string PhoneNumbers = "PhoneNumbers";
        public const string PhoneNumber = "PhoneNumber";
        public const string AssistantName = "AssistantName";
        public const string Birthday = "Birthday";
        public const string BusinessHomePage = "BusinessHomePage";
        public const string Children = "Children";
        public const string Companies = "Companies";
        public const string ContactSource = "ContactSource";
        public const string Department = "Department";
        public const string Generation = "Generation";
        public const string ImAddresses = "ImAddresses";
        public const string ImAddress = "ImAddress";
        public const string JobTitle = "JobTitle";
        public const string Manager = "Manager";
        public const string Mileage = "Mileage";
        public const string OfficeLocation = "OfficeLocation";
        public const string PostalAddressIndex = "PostalAddressIndex";
        public const string Profession = "Profession";
        public const string SpouseName = "SpouseName";
        public const string Surname = "Surname";
        public const string WeddingAnniversary = "WeddingAnniversary";
        public const string HasPicture = "HasPicture";
        public const string Title = "Title";
        public const string FirstName = "FirstName";
        public const string LastName = "LastName";
        public const string Suffix = "Suffix";
        public const string FullName = "FullName";
        public const string YomiFirstName = "YomiFirstName";
        public const string YomiLastName = "YomiLastName";
        public const string Contact = "Contact";
        public const string Entry = "Entry";
        public const string Street = "Street";
        public const string City = "City";
        public const string State = "State";
        public const string SharePointSiteUrl = "SharePointSiteUrl";
        public const string Country = "Country";
        public const string CountryOrRegion = "CountryOrRegion";
        public const string PostalCode = "PostalCode";
        public const string PostOfficeBox = "PostOfficeBox";
        public const string Members = "Members";
        public const string Member = "Member";
        public const string AdditionalProperties = "AdditionalProperties";
        public const string ExtendedFieldURI = "ExtendedFieldURI";
        public const string Value = "Value";
        public const string Values = "Values";
        public const string ToFolderId = "ToFolderId";
        public const string ActualWork = "ActualWork";
        public const string AssignedTime = "AssignedTime";
        public const string BillingInformation = "BillingInformation";
        public const string ChangeCount = "ChangeCount";
        public const string CompleteDate = "CompleteDate";
        public const string Contacts = "Contacts";
        public const string DelegationState = "DelegationState";
        public const string Delegator = "Delegator";
        public const string DueDate = "DueDate";
        public const string IsAssignmentEditable = "IsAssignmentEditable";
        public const string IsComplete = "IsComplete";
        public const string IsTeamTask = "IsTeamTask";
        public const string Owner = "Owner";
        public const string PercentComplete = "PercentComplete";
        public const string Status = "Status";
        public const string StatusDescription = "StatusDescription";
        public const string TotalWork = "TotalWork";
        public const string Task = "Task";
        public const string MailboxCulture = "MailboxCulture";
        public const string MeetingRequestType = "MeetingRequestType";
        public const string IntendedFreeBusyStatus = "IntendedFreeBusyStatus";
        public const string MeetingRequest = "MeetingRequest";
        public const string MeetingResponse = "MeetingResponse";
        public const string MeetingCancellation = "MeetingCancellation";
        public const string ChangeHighlights = "ChangeHighlights";
        public const string HasLocationChanged = "HasLocationChanged";
        public const string HasStartTimeChanged = "HasStartTimeChanged";
        public const string HasEndTimeChanged = "HasEndTimeChanged";
        public const string BaseOffset = "BaseOffset";
        public const string Offset = "Offset";
        public const string Standard = "Standard";
        public const string Daylight = "Daylight";
        public const string Time = "Time";
        public const string AbsoluteDate = "AbsoluteDate";
        public const string UnresolvedEntry = "UnresolvedEntry";
        public const string ResolutionSet = "ResolutionSet";
        public const string Resolution = "Resolution";
        public const string DistributionList = "DistributionList";
        public const string DLExpansion = "DLExpansion";
        public const string IndexedFieldURI = "IndexedFieldURI";
        public const string PullSubscriptionRequest = "PullSubscriptionRequest";
        public const string PushSubscriptionRequest = "PushSubscriptionRequest";
        public const string StreamingSubscriptionRequest = "StreamingSubscriptionRequest";
        public const string EventTypes = "EventTypes";
        public const string EventType = "EventType";
        public const string Timeout = "Timeout";
        public const string Watermark = "Watermark";
        public const string SubscriptionId = "SubscriptionId";
        public const string SubscriptionIds = "SubscriptionIds";
        public const string StatusFrequency = "StatusFrequency";
        public const string URL = "URL";
        public const string CallerData = "CallerData";
        public const string Notification = "Notification";
        public const string Notifications = "Notifications";
        public const string PreviousWatermark = "PreviousWatermark";
        public const string MoreEvents = "MoreEvents";
        public const string TimeStamp = "TimeStamp";
        public const string UnreadCount = "UnreadCount";
        public const string OldParentFolderId = "OldParentFolderId";
        public const string CopiedEvent = "CopiedEvent";
        public const string CreatedEvent = "CreatedEvent";
        public const string DeletedEvent = "DeletedEvent";
        public const string ModifiedEvent = "ModifiedEvent";
        public const string MovedEvent = "MovedEvent";
        public const string NewMailEvent = "NewMailEvent";
        public const string StatusEvent = "StatusEvent";
        public const string FreeBusyChangedEvent = "FreeBusyChangedEvent";
        public const string ExchangeImpersonation = "ExchangeImpersonation";
        public const string ConnectingSID = "ConnectingSID";
        public const string OpenAsAdminOrSystemService = "OpenAsAdminOrSystemService";
        public const string LogonType = "LogonType";
        public const string BudgetType = "BudgetType";
        public const string ManagementRole = "ManagementRole";
        public const string UserRoles = "UserRoles";
        public const string ApplicationRoles = "ApplicationRoles";
        public const string Role = "Role";
        public const string SyncFolderId = "SyncFolderId";
        public const string SyncScope = "SyncScope";
        public const string SyncState = "SyncState";
        public const string Ignore = "Ignore";
        public const string MaxChangesReturned = "MaxChangesReturned";
        public const string Changes = "Changes";
        public const string IncludesLastItemInRange = "IncludesLastItemInRange";
        public const string IncludesLastFolderInRange = "IncludesLastFolderInRange";
        public const string Create = "Create";
        public const string Update = "Update";
        public const string Delete = "Delete";
        public const string ReadFlagChange = "ReadFlagChange";
        public const string SearchParameters = "SearchParameters";
        public const string SoftDeleted = "SoftDeleted";
        public const string Shallow = "Shallow";
        public const string Associated = "Associated";
        public const string BaseFolderId = "BaseFolderId";
        public const string BaseFolderIds = "BaseFolderIds";
        public const string SortOrder = "SortOrder";
        public const string FieldOrder = "FieldOrder";
        public const string CanDelete = "CanDelete";
        public const string CanRenameOrMove = "CanRenameOrMove";
        public const string MustDisplayComment = "MustDisplayComment";
        public const string HasQuota = "HasQuota";
        public const string IsManagedFoldersRoot = "IsManagedFoldersRoot";
        public const string ManagedFolderId = "ManagedFolderId";
        public const string Comment = "Comment";
        public const string StorageQuota = "StorageQuota";
        public const string FolderSize = "FolderSize";
        public const string HomePage = "HomePage";
        public const string ManagedFolderInformation = "ManagedFolderInformation";
        public const string CalendarView = "CalendarView";
        public const string PostedTime = "PostedTime";
        public const string PostItem = "PostItem";
        public const string RequestVersion = "RequestVersion";
        public const string RequestServerVersion = "RequestServerVersion";
        public const string PostReplyItem = "PostReplyItem";
        public const string CreateAssociated = "CreateAssociated";
        public const string CreateContents = "CreateContents";
        public const string CreateHierarchy = "CreateHierarchy";
        public const string Modify = "Modify";
        public const string Read = "Read";
        public const string EffectiveRights = "EffectiveRights";
        public const string LastModifiedName = "LastModifiedName";
        public const string LastModifiedTime = "LastModifiedTime";
        public const string ConversationId = "ConversationId";
        public const string UniqueBody = "UniqueBody";
        public const string BodyType = "BodyType";
        public const string NormalizedBodyType = "NormalizedBodyType";
        public const string UniqueBodyType = "UniqueBodyType";
        public const string AttachmentShape = "AttachmentShape";
        public const string UserId = "UserId";
        public const string UserIds = "UserIds";
        public const string CanCreateItems = "CanCreateItems";
        public const string CanCreateSubFolders = "CanCreateSubFolders";
        public const string IsFolderOwner = "IsFolderOwner";
        public const string IsFolderVisible = "IsFolderVisible";
        public const string IsFolderContact = "IsFolderContact";
        public const string EditItems = "EditItems";
        public const string DeleteItems = "DeleteItems";
        public const string ReadItems = "ReadItems";
        public const string PermissionLevel = "PermissionLevel";
        public const string CalendarPermissionLevel = "CalendarPermissionLevel";
        public const string SID = "SID";
        public const string PrimarySmtpAddress = "PrimarySmtpAddress";
        public const string DistinguishedUser = "DistinguishedUser";
        public const string PermissionSet = "PermissionSet";
        public const string Permissions = "Permissions";
        public const string Permission = "Permission";
        public const string CalendarPermissions = "CalendarPermissions";
        public const string CalendarPermission = "CalendarPermission";
        public const string GroupBy = "GroupBy";
        public const string AggregateOn = "AggregateOn";
        public const string Groups = "Groups";
        public const string GroupedItems = "GroupedItems";
        public const string GroupIndex = "GroupIndex";
        public const string ConflictResults = "ConflictResults";
        public const string Count = "Count";
        public const string OofSettings = "OofSettings";
        public const string UserOofSettings = "UserOofSettings";
        public const string OofState = "OofState";
        public const string ExternalAudience = "ExternalAudience";
        public const string AllowExternalOof = "AllowExternalOof";
        public const string InternalReply = "InternalReply";
        public const string ExternalReply = "ExternalReply";
        public const string Bias = "Bias";
        public const string DayOrder = "DayOrder";
        public const string Year = "Year";
        public const string StandardTime = "StandardTime";
        public const string DaylightTime = "DaylightTime";
        public const string MailboxData = "MailboxData";
        public const string MailboxDataArray = "MailboxDataArray";
        public const string Email = "Email";
        public const string AttendeeType = "AttendeeType";
        public const string ExcludeConflicts = "ExcludeConflicts";
        public const string FreeBusyViewOptions = "FreeBusyViewOptions";
        public const string SuggestionsViewOptions = "SuggestionsViewOptions";
        public const string FreeBusyView = "FreeBusyView";
        public const string TimeWindow = "TimeWindow";
        public const string MergedFreeBusyIntervalInMinutes = "MergedFreeBusyIntervalInMinutes";
        public const string RequestedView = "RequestedView";
        public const string FreeBusyViewType = "FreeBusyViewType";
        public const string CalendarEventArray = "CalendarEventArray";
        public const string CalendarEvent = "CalendarEvent";
        public const string BusyType = "BusyType";
        public const string MergedFreeBusy = "MergedFreeBusy";
        public const string WorkingHours = "WorkingHours";
        public const string WorkingPeriodArray = "WorkingPeriodArray";
        public const string WorkingPeriod = "WorkingPeriod";
        public const string StartTimeInMinutes = "StartTimeInMinutes";
        public const string EndTimeInMinutes = "EndTimeInMinutes";
        public const string GoodThreshold = "GoodThreshold";
        public const string MaximumResultsByDay = "MaximumResultsByDay";
        public const string MaximumNonWorkHourResultsByDay = "MaximumNonWorkHourResultsByDay";
        public const string MeetingDurationInMinutes = "MeetingDurationInMinutes";
        public const string MinimumSuggestionQuality = "MinimumSuggestionQuality";
        public const string DetailedSuggestionsWindow = "DetailedSuggestionsWindow";
        public const string CurrentMeetingTime = "CurrentMeetingTime";
        public const string GlobalObjectId = "GlobalObjectId";
        public const string SuggestionDayResultArray = "SuggestionDayResultArray";
        public const string SuggestionDayResult = "SuggestionDayResult";
        public const string Date = "Date";
        public const string DayQuality = "DayQuality";
        public const string SuggestionArray = "SuggestionArray";
        public const string Suggestion = "Suggestion";
        public const string MeetingTime = "MeetingTime";
        public const string IsWorkTime = "IsWorkTime";
        public const string SuggestionQuality = "SuggestionQuality";
        public const string AttendeeConflictDataArray = "AttendeeConflictDataArray";
        public const string UnknownAttendeeConflictData = "UnknownAttendeeConflictData";
        public const string TooBigGroupAttendeeConflictData = "TooBigGroupAttendeeConflictData";
        public const string IndividualAttendeeConflictData = "IndividualAttendeeConflictData";
        public const string GroupAttendeeConflictData = "GroupAttendeeConflictData";
        public const string NumberOfMembers = "NumberOfMembers";
        public const string NumberOfMembersAvailable = "NumberOfMembersAvailable";
        public const string NumberOfMembersWithConflict = "NumberOfMembersWithConflict";
        public const string NumberOfMembersWithNoData = "NumberOfMembersWithNoData";
        public const string SourceIds = "SourceIds";
        public const string AlternateId = "AlternateId";
        public const string AlternatePublicFolderId = "AlternatePublicFolderId";
        public const string AlternatePublicFolderItemId = "AlternatePublicFolderItemId";
        public const string DelegatePermissions = "DelegatePermissions";
        public const string ReceiveCopiesOfMeetingMessages = "ReceiveCopiesOfMeetingMessages";
        public const string ViewPrivateItems = "ViewPrivateItems";
        public const string CalendarFolderPermissionLevel = "CalendarFolderPermissionLevel";
        public const string TasksFolderPermissionLevel = "TasksFolderPermissionLevel";
        public const string InboxFolderPermissionLevel = "InboxFolderPermissionLevel";
        public const string ContactsFolderPermissionLevel = "ContactsFolderPermissionLevel";
        public const string NotesFolderPermissionLevel = "NotesFolderPermissionLevel";
        public const string JournalFolderPermissionLevel = "JournalFolderPermissionLevel";
        public const string DelegateUser = "DelegateUser";
        public const string DelegateUsers = "DelegateUsers";
        public const string DeliverMeetingRequests = "DeliverMeetingRequests";
        public const string MessageXml = "MessageXml";
        public const string UserConfiguration = "UserConfiguration";
        public const string UserConfigurationName = "UserConfigurationName";
        public const string UserConfigurationProperties = "UserConfigurationProperties";
        public const string Dictionary = "Dictionary";
        public const string DictionaryEntry = "DictionaryEntry";
        public const string DictionaryKey = "DictionaryKey";
        public const string DictionaryValue = "DictionaryValue";
        public const string XmlData = "XmlData";
        public const string BinaryData = "BinaryData";
        public const string FilterHtmlContent = "FilterHtmlContent";
        public const string ConvertHtmlCodePageToUTF8 = "ConvertHtmlCodePageToUTF8";
        public const string UnknownEntries = "UnknownEntries";
        public const string UnknownEntry = "UnknownEntry";
        public const string PasswordExpirationDate = "PasswordExpirationDate";
        public const string Flag = "Flag";
        public const string PersonaPostalAddress = "PostalAddress";
        public const string PostalAddressType = "Type";
        public const string EnhancedLocation = "EnhancedLocation";
        public const string LocationDisplayName = "DisplayName";
        public const string LocationAnnotation = "Annotation";
        public const string LocationSource = "LocationSource";
        public const string LocationUri = "LocationUri";
        public const string Latitude = "Latitude";
        public const string Longitude = "Longitude";
        public const string Accuracy = "Accuracy";
        public const string Altitude = "Altitude";
        public const string AltitudeAccuracy = "AltitudeAccuracy";
        public const string FormattedAddress = "FormattedAddress";
        public const string Guid = "Guid";
        public const string PhoneCallId = "PhoneCallId";
        public const string DialString = "DialString";
        public const string PhoneCallInformation = "PhoneCallInformation";
        public const string PhoneCallState = "PhoneCallState";
        public const string ConnectionFailureCause = "ConnectionFailureCause";
        public const string SIPResponseCode = "SIPResponseCode";
        public const string SIPResponseText = "SIPResponseText";
        public const string WebClientReadFormQueryString = "WebClientReadFormQueryString";
        public const string WebClientEditFormQueryString = "WebClientEditFormQueryString";
        public const string Ids = "Ids";
        public const string Id = "Id";
        public const string TimeZoneDefinitions = "TimeZoneDefinitions";
        public const string TimeZoneDefinition = "TimeZoneDefinition";
        public const string Periods = "Periods";
        public const string Period = "Period";
        public const string TransitionsGroups = "TransitionsGroups";
        public const string TransitionsGroup = "TransitionsGroup";
        public const string Transitions = "Transitions";
        public const string Transition = "Transition";
        public const string AbsoluteDateTransition = "AbsoluteDateTransition";
        public const string RecurringDayTransition = "RecurringDayTransition";
        public const string RecurringDateTransition = "RecurringDateTransition";
        public const string DateTime = "DateTime";
        public const string TimeOffset = "TimeOffset";
        public const string Day = "Day";
        public const string TimeZoneContext = "TimeZoneContext";
        public const string StartTimeZone = "StartTimeZone";
        public const string EndTimeZone = "EndTimeZone";
        public const string ReceivedBy = "ReceivedBy";
        public const string ReceivedRepresenting = "ReceivedRepresenting";
        public const string Uid = "UID";
        public const string RecurrenceId = "RecurrenceId";
        public const string DateTimeStamp = "DateTimeStamp";
        public const string IsInline = "IsInline";
        public const string IsContactPhoto = "IsContactPhoto";
        public const string QueryString = "QueryString";
        public const string HighlightTerms = "HighlightTerms";
        public const string HighlightTerm = "Term";
        public const string HighlightTermScope = "Scope";
        public const string HighlightTermValue = "Value";
        public const string CalendarEventDetails = "CalendarEventDetails";
        public const string ID = "ID";
        public const string IsException = "IsException";
        public const string IsReminderSet = "IsReminderSet";
        public const string IsPrivate = "IsPrivate";
        public const string FirstDayOfWeek = "FirstDayOfWeek";
        public const string Verb = "Verb";
        public const string Parameter = "Parameter";
        public const string ReturnValue = "ReturnValue";
        public const string ReturnNewItemIds = "ReturnNewItemIds";
        public const string DateTimePrecision = "DateTimePrecision";
        public const string ConvertInlineImagesToDataUrls = "ConvertInlineImagesToDataUrls";
        public const string InlineImageUrlTemplate = "InlineImageUrlTemplate";
        public const string BlockExternalImages = "BlockExternalImages";
        public const string AddBlankTargetToLinks = "AddBlankTargetToLinks";
        public const string MaximumBodySize = "MaximumBodySize";
        public const string StoreEntryId = "StoreEntryId";
        public const string InstanceKey = "InstanceKey";
        public const string NormalizedBody = "NormalizedBody";
        public const string PolicyTag = "PolicyTag";
        public const string ArchiveTag = "ArchiveTag";
        public const string RetentionDate = "RetentionDate";
        public const string DisableReason = "DisableReason";
        public const string AppMarketplaceUrl = "AppMarketplaceUrl";
        public const string TextBody = "TextBody";
        public const string IconIndex = "IconIndex";
        public const string GlobalIconIndex = "GlobalIconIndex";
        public const string DraftItemIds = "DraftItemIds";
        public const string HasIrm = "HasIrm";
        public const string GlobalHasIrm = "GlobalHasIrm";
        public const string ApprovalRequestData = "ApprovalRequestData";
        public const string IsUndecidedApprovalRequest = "IsUndecidedApprovalRequest";
        public const string ApprovalDecision = "ApprovalDecision";
        public const string ApprovalDecisionMaker = "ApprovalDecisionMaker";
        public const string ApprovalDecisionTime = "ApprovalDecisionTime";
        public const string VotingOptionData = "VotingOptionData";
        public const string VotingOptionDisplayName = "DisplayName";
        public const string SendPrompt = "SendPrompt";
        public const string VotingInformation = "VotingInformation";
        public const string UserOptions = "UserOptions";
        public const string VotingResponse = "VotingResponse";
        public const string NumberOfDays = "NumberOfDays";
        public const string AcceptanceState = "AcceptanceState";

        public const string NlgEntityExtractionResult = "EntityExtractionResult";
        public const string NlgAddresses = "Addresses";
        public const string NlgAddress = "Address";
        public const string NlgMeetingSuggestions = "MeetingSuggestions";
        public const string NlgMeetingSuggestion = "MeetingSuggestion";
        public const string NlgTaskSuggestions = "TaskSuggestions";
        public const string NlgTaskSuggestion = "TaskSuggestion";
        public const string NlgBusinessName = "BusinessName";
        public const string NlgPeopleName = "PeopleName";
        public const string NlgEmailAddresses = "EmailAddresses";
        public const string NlgEmailAddress = "EmailAddress";
        public const string NlgEmailPosition = "Position";
        public const string NlgContacts = "Contacts";
        public const string NlgContact = "Contact";
        public const string NlgContactString = "ContactString";
        public const string NlgUrls = "Urls";
        public const string NlgUrl = "Url";
        public const string NlgPhoneNumbers = "PhoneNumbers";
        public const string NlgPhone = "Phone";
        public const string NlgAttendees = "Attendees";
        public const string NlgEmailUser = "EmailUser";
        public const string NlgLocation = "Location";
        public const string NlgSubject = "Subject";
        public const string NlgMeetingString = "MeetingString";
        public const string NlgStartTime = "StartTime";
        public const string NlgEndTime = "EndTime";
        public const string NlgTaskString = "TaskString";
        public const string NlgAssignees = "Assignees";
        public const string NlgPersonName = "PersonName";
        public const string NlgOriginalPhoneString = "OriginalPhoneString";
        public const string NlgPhoneString = "PhoneString";
        public const string NlgType = "Type";
        public const string NlgName = "Name";
        public const string NlgUserId = "UserId";

        public const string GetClientAccessToken = "GetClientAccessToken";
        public const string GetClientAccessTokenResponse = "GetClientAccessTokenResponse";
        public const string GetClientAccessTokenResponseMessage = "GetClientAccessTokenResponseMessage";
        public const string TokenRequests = "TokenRequests";
        public const string TokenRequest = "TokenRequest";
        public const string TokenType = "TokenType";
        public const string TokenValue = "TokenValue";
        public const string TTL = "TTL";
        public const string Tokens = "Tokens";

        public const string MarkAsJunk = "MarkAsJunk";
        public const string MarkAsJunkResponse = "MarkAsJunkResponse";
        public const string MarkAsJunkResponseMessage = "MarkAsJunkResponseMessage";
        public const string MovedItemId = "MovedItemId";

        #region Persona

        public const string CreationTime = "CreationTime";
        public const string People = "People";
        public const string Persona = "Persona";
        public const string PersonaId = "PersonaId";
        public const string PersonaShape = "PersonaShape";
        public const string RelevanceScore = "RelevanceScore";
        public const string TotalNumberOfPeopleInView = "TotalNumberOfPeopleInView";
        public const string FirstMatchingRowIndex = "FirstMatchingRowIndex";
        public const string FirstLoadedRowIndex = "FirstLoadedRowIndex";
        public const string YomiCompanyName = "YomiCompanyName";
        public const string Emails1 = "Emails1";
        public const string Emails2 = "Emails2";
        public const string Emails3 = "Emails3";
        public const string HomeAddresses = "HomeAddresses";
        public const string BusinessAddresses = "BusinessAddresses";
        public const string OtherAddresses = "OtherAddresses";
        public const string BusinessPhoneNumbers = "BusinessPhoneNumbers";
        public const string BusinessPhoneNumbers2 = "BusinessPhoneNumbers2";
        public const string AssistantPhoneNumbers = "AssistantPhoneNumbers";
        public const string TTYTDDPhoneNumbers = "TTYTDDPhoneNumbers";
        public const string HomePhones = "HomePhones";
        public const string HomePhones2 = "HomePhones2";
        public const string MobilePhones = "MobilePhones";
        public const string MobilePhones2 = "MobilePhones2";
        public const string CallbackPhones = "CallbackPhones";
        public const string CarPhones = "CarPhones";
        public const string HomeFaxes = "HomeFaxes";
        public const string OrganizationMainPhones = "OrganizationMainPhones";
        public const string OtherFaxes = "OtherFaxes";
        public const string OtherTelephones = "OtherTelephones";
        public const string OtherPhones2 = "OtherPhones2";
        public const string Pagers = "Pagers";
        public const string RadioPhones = "RadioPhones";
        public const string TelexNumbers = "TelexNumbers";
        public const string WorkFaxes = "WorkFaxes";
        public const string FileAses = "FileAses";
        public const string CompanyNames = "CompanyNames";
        public const string DisplayNames = "DisplayNames";
        public const string DisplayNamePrefixes = "DisplayNamePrefixes";
        public const string GivenNames = "GivenNames";
        public const string MiddleNames = "MiddleNames";
        public const string Surnames = "Surnames";
        public const string Generations = "Generations";
        public const string Nicknames = "Nicknames";
        public const string YomiCompanyNames = "YomiCompanyNames";
        public const string YomiFirstNames = "YomiFirstNames";
        public const string YomiLastNames = "YomiLastNames";
        public const string Managers = "Managers";
        public const string AssistantNames = "AssistantNames";
        public const string Professions = "Professions";
        public const string SpouseNames = "SpouseNames";
        public const string Departments = "Departments";
        public const string Titles = "Titles";
        public const string ImAddresses2 = "ImAddresses2";
        public const string ImAddresses3 = "ImAddresses3";
        public const string DisplayNamePrefix = "DisplayNamePrefix";
        public const string DisplayNameFirstLast = "DisplayNameFirstLast";
        public const string DisplayNameLastFirst = "DisplayNameLastFirst";
        public const string DisplayNameFirstLastHeader = "DisplayNameFirstLastHeader";
        public const string DisplayNameLastFirstHeader = "DisplayNameLastFirstHeader";
        public const string IsFavorite = "IsFavorite";
        public const string Schools = "Schools";
        public const string Hobbies = "Hobbies";
        public const string Locations = "Locations";
        public const string OfficeLocations = "OfficeLocations";
        public const string BusinessHomePages = "BusinessHomePages";
        public const string PersonalHomePages = "PersonalHomePages";
        public const string ThirdPartyPhotoUrls = "ThirdPartyPhotoUrls";
        public const string Attribution = "Attribution";
        public const string Attributions = "Attributions";
        public const string StringAttributedValue = "StringAttributedValue";
        public const string DisplayNameFirstLastSortKey = "DisplayNameFirstLastSortKey";
        public const string DisplayNameLastFirstSortKey = "DisplayNameLastFirstSortKey";
        public const string CompanyNameSortKey = "CompanyNameSortKey";
        public const string HomeCitySortKey = "HomeCitySortKey";
        public const string WorkCitySortKey = "WorkCitySortKey";
        public const string FileAsId = "FileAsId";
        public const string FileAsIds = "FileAsIds";
        public const string HomeCity = "HomeCity";
        public const string WorkCity = "WorkCity";
        public const string PersonaType = "PersonaType";
        public const string Birthdays = "Birthdays";
        public const string BirthdaysLocal = "BirthdaysLocal";
        public const string WeddingAnniversaries = "WeddingAnniversaries";
        public const string WeddingAnniversariesLocal = "WeddingAnniversariesLocal";
        public const string OriginalDisplayName = "OriginalDisplayName";

        #endregion

        #region People Insights
        public const string Person = "Person";
        public const string Insights = "Insights";
        public const string Insight = "Insight";
        public const string InsightType = "InsightType";
        public const string InsightSourceType = "InsightSourceType";
        public const string InsightValue = "InsightValue";
        public const string InsightSource = "InsightSource";
        public const string UpdatedUtcTicks = "UpdatedUtcTicks";
        public const string StringInsightValue = "StringInsightValue";
        public const string ProfileInsightValue = "ProfileInsightValue";
        public const string JobInsightValue = "JobInsightValue";
        public const string OutOfOfficeInsightValue = "OutOfOfficeInsightValue";
        public const string UserProfilePicture = "UserProfilePicture";
        public const string EducationInsightValue = "EducationInsightValue";
        public const string SkillInsightValue = "SkillInsightValue";
        public const string MeetingInsightValue = "MeetingInsightValue";
        public const string Attendees = "Attendees";
        public const string EmailInsightValue = "EmailInsightValue";
        public const string ThreadId = "ThreadId";
        public const string LastEmailDateUtcTicks = "LastEmailDateUtcTicks";
        public const string LastEmailSender = "LastEmailSender";
        public const string EmailsCount = "EmailsCount";
        public const string DelveDocument = "DelveDocument";
        public const string CompanyInsightValue = "CompanyInsightValue";
        public const string ArrayOfInsightValue = "ArrayOfInsightValue";
        public const string InsightContent = "InsightContent";
        public const string SingleValueInsightContent = "SingleValueInsightContent";
        public const string MultiValueInsightContent = "MultiValueInsightContent";
        public const string ArrayOfInsight = "ArrayOfInsight";
        public const string PersonType = "PersonType";
        public const string SatoriId = "SatoriId";
        public const string DescriptionAttribution = "DescriptionAttribution";
        public const string ImageUrl = "ImageUrl";
        public const string ImageUrlAttribution = "ImageUrlAttribution";
        public const string YearFound = "YearFound";
        public const string FinanceSymbol = "FinanceSymbol";
        public const string WebsiteUrl = "WebsiteUrl";
        public const string Rank = "Rank";
        public const string Author = "Author";
        public const string Created = "Created";
        public const string DefaultEncodingURL = "DefaultEncodingURL";
        public const string FileType = "FileType";
        public const string Data = "Data";
        public const string ItemList = "ItemList";
        public const string Avatar = "Avatar";
        public const string JoinedUtcTicks = "JoinedUtcTicks";
        public const string Company = "Company";
        public const string StartUtcTicks = "StartUtcTicks";
        public const string EndUtcTicks = "EndUtcTicks";
        public const string Blob = "Blob";
        public const string PhotoSize = "PhotoSize";
        public const string Institute = "Institute";
        public const string Degree = "Degree";
        public const string Strength = "Strength";
        public const string ComputedInsightValueProperty = "ComputedInsightValueProperty";
        public const string ComputedInsightValue = "ComputedInsightValue";
        public const string Properties = "Properties";
        public const string Property = "Property";
        public const string Key = "Key";
        public const string SMSNumber = "SMSNumber";
        public const string FacebookProfileLink = "FacebookProfileLink";
        public const string LinkedInProfileLink = "LinkedInProfileLink";
        public const string ProfessionalBiography = "ProfessionalBiography";
        public const string TeamSize = "TeamSize";
        public const string Hometown = "Hometown";
        public const string CurrentLocation = "CurrentLocation";
        public const string Office = "Office";
        public const string Headline = "Headline";
        public const string ManagementChain = "ManagementChain";
        public const string Peers = "Peers";
        public const string MutualConnections = "MutualConnections";
        public const string MutualManager = "MutualManager";
        public const string Skills = "Skills";
        public const string JobInsight = "JobInsight";
        public const string CurrentJob = "CurrentJob";
        public const string CompanyProfile = "CompanyProfile";
        public const string CompanyInsight = "CompanyInsight";
        public const string Text = "Text";
        public const string ImageType = "ImageType";
        public const string DocumentId = "DocumentId";
        public const string PreviewURL = "PreviewURL";
        public const string LastEditor = "LastEditor";
        public const string ProfilePicture = "ProfilePicture";

        #endregion

        #region Conversations

        public const string Conversations = "Conversations";
        public const string Conversation = "Conversation";
        public const string UniqueRecipients = "UniqueRecipients";
        public const string GlobalUniqueRecipients = "GlobalUniqueRecipients";
        public const string UniqueUnreadSenders = "UniqueUnreadSenders";
        public const string GlobalUniqueUnreadSenders = "GlobalUniqueUnreadSenders";
        public const string UniqueSenders = "UniqueSenders";
        public const string GlobalUniqueSenders = "GlobalUniqueSenders";
        public const string LastDeliveryTime = "LastDeliveryTime";
        public const string GlobalLastDeliveryTime = "GlobalLastDeliveryTime";
        public const string GlobalCategories = "GlobalCategories";
        public const string FlagStatus = "FlagStatus";
        public const string GlobalFlagStatus = "GlobalFlagStatus";
        public const string GlobalHasAttachments = "GlobalHasAttachments";
        public const string MessageCount = "MessageCount";
        public const string GlobalMessageCount = "GlobalMessageCount";
        public const string GlobalUnreadCount = "GlobalUnreadCount";
        public const string GlobalSize = "GlobalSize";
        public const string ItemClasses = "ItemClasses";
        public const string GlobalItemClasses = "GlobalItemClasses";
        public const string GlobalImportance = "GlobalImportance";
        public const string GlobalInferredImportance = "GlobalInferredImportance";
        public const string GlobalItemIds = "GlobalItemIds";
        public const string ChangeType = "ChangeType";
        public const string ReadFlag = "ReadFlag";
        public const string TotalConversationsInView = "TotalConversationsInView";
        public const string IndexedOffset = "IndexedOffset";
        public const string ConversationShape = "ConversationShape";
        public const string MailboxScope = "MailboxScope";

        // ApplyConversationAction
        public const string ApplyConversationAction = "ApplyConversationAction";
        public const string ConversationActions = "ConversationActions";
        public const string ConversationAction = "ConversationAction";
        public const string ApplyConversationActionResponse = "ApplyConversationActionResponse";
        public const string ApplyConversationActionResponseMessage = "ApplyConversationActionResponseMessage";
        public const string EnableAlwaysDelete = "EnableAlwaysDelete";
        public const string ProcessRightAway = "ProcessRightAway";
        public const string DestinationFolderId = "DestinationFolderId";
        public const string ContextFolderId = "ContextFolderId";
        public const string ConversationLastSyncTime = "ConversationLastSyncTime";
        public const string AlwaysCategorize = "AlwaysCategorize";
        public const string AlwaysDelete = "AlwaysDelete";
        public const string AlwaysMove = "AlwaysMove";
        public const string Move = "Move";
        public const string Copy = "Copy";
        public const string SetReadState = "SetReadState";
        public const string SetRetentionPolicy = "SetRetentionPolicy";
        public const string DeleteType = "DeleteType";
        public const string RetentionPolicyType = "RetentionPolicyType";
        public const string RetentionPolicyTagId = "RetentionPolicyTagId";

        // GetConversationItems
        public const string FoldersToIgnore = "FoldersToIgnore";
        public const string ParentInternetMessageId = "ParentInternetMessageId";
        public const string ConversationNode = "ConversationNode";
        public const string ConversationNodes = "ConversationNodes";
        public const string MaxItemsToReturn = "MaxItemsToReturn";

        #endregion

        #region TeamMailbox

        public const string SetTeamMailbox = "SetTeamMailbox";
        public const string SetTeamMailboxResponse = "SetTeamMailboxResponse";
        public const string UnpinTeamMailbox = "UnpinTeamMailbox";
        public const string UnpinTeamMailboxResponse = "UnpinTeamMailboxResponse";

        #endregion

        #region RoomList & Room

        public const string RoomLists = "RoomLists";
        public const string Rooms = "Rooms";
        public const string Room = "Room";
        public const string RoomList = "RoomList";
        public const string RoomId = "Id";

        #endregion

        #region Autodiscover

        public const string Autodiscover = "Autodiscover";
        public const string BinarySecret = "BinarySecret";
        public const string Response = "Response";
        public const string User = "User";
        public const string LegacyDN = "LegacyDN";
        public const string DeploymentId = "DeploymentId";
        public const string Account = "Account";
        public const string AccountType = "AccountType";
        public const string Action = "Action";
        public const string To = "To";
        public const string RedirectAddr = "RedirectAddr";
        public const string RedirectUrl = "RedirectUrl";
        public const string Protocol = "Protocol";
        public const string Type = "Type";
        public const string Server = "Server";
        public const string OwnerSmtpAddress = "OwnerSmtpAddress";
        public const string ServerDN = "ServerDN";
        public const string ServerVersion = "ServerVersion";
        public const string ServerVersionInfo = "ServerVersionInfo";
        public const string AD = "AD";
        public const string AuthPackage = "AuthPackage";
        public const string MdbDN = "MdbDN";
        public const string EWSUrl = "EwsUrl"; // Server side emits "Ews", not "EWS".
        public const string EwsPartnerUrl = "EwsPartnerUrl";
        public const string EmwsUrl = "EmwsUrl";
        public const string ASUrl = "ASUrl";
        public const string OOFUrl = "OOFUrl";
        public const string UMUrl = "UMUrl";
        public const string OABUrl = "OABUrl";
        public const string Internal = "Internal";
        public const string External = "External";
        public const string OWAUrl = "OWAUrl";
        public const string Error = "Error";
        public const string ErrorCode = "ErrorCode";
        public const string DebugData = "DebugData";
        public const string Users = "Users";
        public const string RequestedSettings = "RequestedSettings";
        public const string Setting = "Setting";
        public const string GetUserSettingsRequestMessage = "GetUserSettingsRequestMessage";
        public const string RequestedServerVersion = "RequestedServerVersion";
        public const string Request = "Request";
        public const string RedirectTarget = "RedirectTarget";
        public const string UserSettings = "UserSettings";
        public const string UserSettingErrors = "UserSettingErrors";
        public const string GetUserSettingsResponseMessage = "GetUserSettingsResponseMessage";
        public const string ErrorMessage = "ErrorMessage";
        public const string UserResponse = "UserResponse";
        public const string UserResponses = "UserResponses";
        public const string UserSettingError = "UserSettingError";
        public const string Domain = "Domain";
        public const string Domains = "Domains";
        public const string DomainResponse = "DomainResponse";
        public const string DomainResponses = "DomainResponses";
        public const string DomainSetting = "DomainSetting";
        public const string DomainSettings = "DomainSettings";
        public const string DomainStringSetting = "DomainStringSetting";
        public const string DomainSettingError = "DomainSettingError";
        public const string DomainSettingErrors = "DomainSettingErrors";
        public const string GetDomainSettingsRequestMessage = "GetDomainSettingsRequestMessage";
        public const string GetDomainSettingsResponseMessage = "GetDomainSettingsResponseMessage";
        public const string SettingName = "SettingName";
        public const string UserSetting = "UserSetting";
        public const string StringSetting = "StringSetting";
        public const string WebClientUrlCollectionSetting = "WebClientUrlCollectionSetting";
        public const string WebClientUrls = "WebClientUrls";
        public const string WebClientUrl = "WebClientUrl";
        public const string AuthenticationMethods = "AuthenticationMethods";
        public const string Url = "Url";
        public const string AlternateMailboxCollectionSetting = "AlternateMailboxCollectionSetting";
        public const string AlternateMailboxes = "AlternateMailboxes";
        public const string AlternateMailbox = "AlternateMailbox";
        public const string ProtocolConnectionCollectionSetting = "ProtocolConnectionCollectionSetting";
        public const string ProtocolConnections = "ProtocolConnections";
        public const string ProtocolConnection = "ProtocolConnection";
        public const string DocumentSharingLocationCollectionSetting = "DocumentSharingLocationCollectionSetting";
        public const string DocumentSharingLocations = "DocumentSharingLocations";
        public const string DocumentSharingLocation = "DocumentSharingLocation";
        public const string ServiceUrl = "ServiceUrl";
        public const string LocationUrl = "LocationUrl";
        public const string SupportedFileExtensions = "SupportedFileExtensions";
        public const string FileExtension = "FileExtension";
        public const string ExternalAccessAllowed = "ExternalAccessAllowed";
        public const string AnonymousAccessAllowed = "AnonymousAccessAllowed";
        public const string CanModifyPermissions = "CanModifyPermissions";
        public const string IsDefault = "IsDefault";
        public const string EncryptionMethod = "EncryptionMethod";
        public const string Hostname = "Hostname";
        public const string Port = "Port";
        public const string Version = "Version";
        public const string MajorVersion = "MajorVersion";
        public const string MinorVersion = "MinorVersion";
        public const string MajorBuildNumber = "MajorBuildNumber";
        public const string MinorBuildNumber = "MinorBuildNumber";
        public const string RequestedVersion = "RequestedVersion";
        public const string PublicFolderServer = "PublicFolderServer";
        public const string Ssl = "SSL";
        public const string SharingUrl = "SharingUrl";
        public const string EcpUrl = "EcpUrl";
        public const string EcpUrl_um = "EcpUrl-um";
        public const string EcpUrl_aggr = "EcpUrl-aggr";
        public const string EcpUrl_sms = "EcpUrl-sms";
        public const string EcpUrl_mt = "EcpUrl-mt";
        public const string EcpUrl_ret = "EcpUrl-ret";
        public const string EcpUrl_publish = "EcpUrl-publish";
        public const string EcpUrl_photo = "EcpUrl-photo";
        public const string ExchangeRpcUrl = "ExchangeRpcUrl";
        public const string EcpUrl_connect = "EcpUrl-connect";
        public const string EcpUrl_tm = "EcpUrl-tm";
        public const string EcpUrl_tmCreating = "EcpUrl-tmCreating";
        public const string EcpUrl_tmEditing = "EcpUrl-tmEditing";
        public const string EcpUrl_tmHiding = "EcpUrl-tmHiding";
        public const string SiteMailboxCreationURL = "SiteMailboxCreationURL";
        public const string EcpUrl_extinstall = "EcpUrl-extinstall";
        public const string PartnerToken = "PartnerToken";
        public const string PartnerTokenReference = "PartnerTokenReference";
        public const string ServerExclusiveConnect = "ServerExclusiveConnect";
        public const string AutoDiscoverSMTPAddress = "AutoDiscoverSMTPAddress";
        public const string CertPrincipalName = "CertPrincipalName";
        public const string GroupingInformation = "GroupingInformation";
        #endregion

        #region InboxRule
        public const string MailboxSmtpAddress = "MailboxSmtpAddress";
        public const string RuleId = "RuleId";
        public const string Priority = "Priority";
        public const string IsEnabled = "IsEnabled";
        public const string IsNotSupported = "IsNotSupported";
        public const string IsInError = "IsInError";
        public const string Conditions = "Conditions";
        public const string Exceptions = "Exceptions";
        public const string Actions = "Actions";
        public const string InboxRules = "InboxRules";
        public const string Rule = "Rule";
        public const string OutlookRuleBlobExists = "OutlookRuleBlobExists";
        public const string RemoveOutlookRuleBlob = "RemoveOutlookRuleBlob";
        public const string ContainsBodyStrings = "ContainsBodyStrings";
        public const string ContainsHeaderStrings = "ContainsHeaderStrings";
        public const string ContainsRecipientStrings = "ContainsRecipientStrings";
        public const string ContainsSenderStrings = "ContainsSenderStrings";
        public const string ContainsSubjectOrBodyStrings = "ContainsSubjectOrBodyStrings";
        public const string ContainsSubjectStrings = "ContainsSubjectStrings";
        public const string FlaggedForAction = "FlaggedForAction";
        public const string FromAddresses = "FromAddresses";
        public const string FromConnectedAccounts = "FromConnectedAccounts";
        public const string IsApprovalRequest = "IsApprovalRequest";
        public const string IsAutomaticForward = "IsAutomaticForward";
        public const string IsAutomaticReply = "IsAutomaticReply";
        public const string IsEncrypted = "IsEncrypted";
        public const string IsMeetingRequest = "IsMeetingRequest";
        public const string IsMeetingResponse = "IsMeetingResponse";
        public const string IsNDR = "IsNDR";
        public const string IsPermissionControlled = "IsPermissionControlled";
        public const string IsSigned = "IsSigned";
        public const string IsVoicemail = "IsVoicemail";
        public const string IsReadReceipt = "IsReadReceipt";
        public const string MessageClassifications = "MessageClassifications";
        public const string NotSentToMe = "NotSentToMe";
        public const string SentCcMe = "SentCcMe";
        public const string SentOnlyToMe = "SentOnlyToMe";
        public const string SentToAddresses = "SentToAddresses";
        public const string SentToMe = "SentToMe";
        public const string SentToOrCcMe = "SentToOrCcMe";
        public const string WithinDateRange = "WithinDateRange";
        public const string WithinSizeRange = "WithinSizeRange";
        public const string MinimumSize = "MinimumSize";
        public const string MaximumSize = "MaximumSize";
        public const string StartDateTime = "StartDateTime";
        public const string EndDateTime = "EndDateTime";
        public const string AssignCategories = "AssignCategories";
        public const string CopyToFolder = "CopyToFolder";
        public const string FlagMessage = "FlagMessage";
        public const string ForwardAsAttachmentToRecipients = "ForwardAsAttachmentToRecipients";
        public const string ForwardToRecipients = "ForwardToRecipients";
        public const string MarkImportance = "MarkImportance";
        public const string MarkAsRead = "MarkAsRead";
        public const string MoveToFolder = "MoveToFolder";
        public const string PermanentDelete = "PermanentDelete";
        public const string RedirectToRecipients = "RedirectToRecipients";
        public const string SendSMSAlertToRecipients = "SendSMSAlertToRecipients";
        public const string ServerReplyWithMessage = "ServerReplyWithMessage";
        public const string StopProcessingRules = "StopProcessingRules";
        public const string CreateRuleOperation = "CreateRuleOperation";
        public const string SetRuleOperation = "SetRuleOperation";
        public const string DeleteRuleOperation = "DeleteRuleOperation";
        public const string Operations = "Operations";
        public const string RuleOperationErrors = "RuleOperationErrors";
        public const string RuleOperationError = "RuleOperationError";
        public const string OperationIndex = "OperationIndex";
        public const string ValidationErrors = "ValidationErrors";
        public const string FieldValue = "FieldValue";
        #endregion

        #region Restrictions
        public const string Not = "Not";
        public const string Bitmask = "Bitmask";
        public const string Constant = "Constant";
        public const string Restriction = "Restriction";
        public const string Condition = "Condition";
        public const string Contains = "Contains";
        public const string Excludes = "Excludes";
        public const string Exists = "Exists";
        public const string FieldURIOrConstant = "FieldURIOrConstant";
        public const string And = "And";
        public const string Or = "Or";
        public const string IsEqualTo = "IsEqualTo";
        public const string IsNotEqualTo = "IsNotEqualTo";
        public const string IsGreaterThan = "IsGreaterThan";
        public const string IsGreaterThanOrEqualTo = "IsGreaterThanOrEqualTo";
        public const string IsLessThan = "IsLessThan";
        public const string IsLessThanOrEqualTo = "IsLessThanOrEqualTo";
        #endregion

        #region Directory only contact properties
        public const string PhoneticFullName = "PhoneticFullName";
        public const string PhoneticFirstName = "PhoneticFirstName";
        public const string PhoneticLastName = "PhoneticLastName";
        public const string Alias = "Alias";
        public const string Notes = "Notes";
        public const string Photo = "Photo";
        public const string UserSMIMECertificate = "UserSMIMECertificate";
        public const string MSExchangeCertificate = "MSExchangeCertificate";
        public const string DirectoryId = "DirectoryId";
        public const string ManagerMailbox = "ManagerMailbox";
        public const string DirectReports = "DirectReports";
        #endregion

        #region Photos

        public const string SizeRequested = "SizeRequested";
        public const string HasChanged = "HasChanged";
        public const string PictureData = "PictureData";

        #endregion

        #region Request/response element names
        public const string ResponseMessage = "ResponseMessage";
        public const string ResponseMessages = "ResponseMessages";

        // FindConversation
        public const string FindConversation = "FindConversation";
        public const string FindConversationResponse = "FindConversationResponse";
        public const string FindConversationResponseMessage = "FindConversationResponseMessage";

        // GetConversationItems
        public const string GetConversationItems = "GetConversationItems";
        public const string GetConversationItemsResponse = "GetConversationItemsResponse";
        public const string GetConversationItemsResponseMessage = "GetConversationItemsResponseMessage";

        // FindItem
        public const string FindItem = "FindItem";
        public const string FindItemResponse = "FindItemResponse";
        public const string FindItemResponseMessage = "FindItemResponseMessage";

        // GetItem
        public const string GetItem = "GetItem";
        public const string GetItemResponse = "GetItemResponse";
        public const string GetItemResponseMessage = "GetItemResponseMessage";

        // CreateItem
        public const string CreateItem = "CreateItem";
        public const string CreateItemResponse = "CreateItemResponse";
        public const string CreateItemResponseMessage = "CreateItemResponseMessage";

        // SendItem
        public const string SendItem = "SendItem";
        public const string SendItemResponse = "SendItemResponse";
        public const string SendItemResponseMessage = "SendItemResponseMessage";

        // DeleteItem
        public const string DeleteItem = "DeleteItem";
        public const string DeleteItemResponse = "DeleteItemResponse";
        public const string DeleteItemResponseMessage = "DeleteItemResponseMessage";

        // UpdateItem
        public const string UpdateItem = "UpdateItem";
        public const string UpdateItemResponse = "UpdateItemResponse";
        public const string UpdateItemResponseMessage = "UpdateItemResponseMessage";

        // CopyItem
        public const string CopyItem = "CopyItem";
        public const string CopyItemResponse = "CopyItemResponse";
        public const string CopyItemResponseMessage = "CopyItemResponseMessage";

        // MoveItem
        public const string MoveItem = "MoveItem";
        public const string MoveItemResponse = "MoveItemResponse";
        public const string MoveItemResponseMessage = "MoveItemResponseMessage";

        // ArchiveItem
        public const string ArchiveItem = "ArchiveItem";
        public const string ArchiveItemResponse = "ArchiveItemResponse";
        public const string ArchiveItemResponseMessage = "ArchiveItemResponseMessage";
        public const string ArchiveSourceFolderId = "ArchiveSourceFolderId";

        // FindFolder
        public const string FindFolder = "FindFolder";
        public const string FindFolderResponse = "FindFolderResponse";
        public const string FindFolderResponseMessage = "FindFolderResponseMessage";

        // GetFolder
        public const string GetFolder = "GetFolder";
        public const string GetFolderResponse = "GetFolderResponse";
        public const string GetFolderResponseMessage = "GetFolderResponseMessage";

        // CreateFolder
        public const string CreateFolder = "CreateFolder";
        public const string CreateFolderResponse = "CreateFolderResponse";
        public const string CreateFolderResponseMessage = "CreateFolderResponseMessage";

        // DeleteFolder
        public const string DeleteFolder = "DeleteFolder";
        public const string DeleteFolderResponse = "DeleteFolderResponse";
        public const string DeleteFolderResponseMessage = "DeleteFolderResponseMessage";

        // EmptyFolder
        public const string EmptyFolder = "EmptyFolder";
        public const string EmptyFolderResponse = "EmptyFolderResponse";
        public const string EmptyFolderResponseMessage = "EmptyFolderResponseMessage";

        // UpdateFolder
        public const string UpdateFolder = "UpdateFolder";
        public const string UpdateFolderResponse = "UpdateFolderResponse";
        public const string UpdateFolderResponseMessage = "UpdateFolderResponseMessage";

        // CopyFolder
        public const string CopyFolder = "CopyFolder";
        public const string CopyFolderResponse = "CopyFolderResponse";
        public const string CopyFolderResponseMessage = "CopyFolderResponseMessage";

        // MoveFolder
        public const string MoveFolder = "MoveFolder";
        public const string MoveFolderResponse = "MoveFolderResponse";
        public const string MoveFolderResponseMessage = "MoveFolderResponseMessage";

        // MarkAllItemsAsRead
        public const string MarkAllItemsAsRead = "MarkAllItemsAsRead";
        public const string MarkAllItemsAsReadResponse = "MarkAllItemsAsReadResponse";
        public const string MarkAllItemsAsReadResponseMessage = "MarkAllItemsAsReadResponseMessage";

        // FindPeople
        public const string FindPeople = "FindPeople";
        public const string FindPeopleResponse = "FindPeopleResponse";
        public const string FindPeopleResponseMessage = "FindPeopleResponseMessage";
        public const string SearchPeopleSuggestionIndex = "SearchPeopleSuggestionIndex";
        public const string SearchPeopleContext = "Context";
        public const string SearchPeopleQuerySources = "QuerySources";
        public const string FindPeopleTransactionId = "TransactionId";
        public const string FindPeopleSources = "Sources";

        // GetPeopleInsights
        public const string GetPeopleInsights = "GetPeopleInsights";
        public const string GetPeopleInsightsResponse = "GetPeopleInsightsResponse";
        public const string GetPeopleInsightsResponseMessage = "GetPeopleInsightsResponseMessage";

        // GetUserPhoto
        public const string GetUserPhoto = "GetUserPhoto";
        public const string GetUserPhotoResponse = "GetUserPhotoResponse";
        public const string GetUserPhotoResponseMessage = "GetUserPhotoResponseMessage";

        // SetUserPhoto
        public const string SetUserPhoto = "SetUserPhoto";
        public const string SetUserPhotoResponse = "SetUserPhotoResponse";
        public const string SetUserPhotoResponseMessage = "SetUserPhotoResponseMessage";


        // GetAttachment
        public const string GetAttachment = "GetAttachment";
        public const string GetAttachmentResponse = "GetAttachmentResponse";
        public const string GetAttachmentResponseMessage = "GetAttachmentResponseMessage";

        // CreateAttachment
        public const string CreateAttachment = "CreateAttachment";
        public const string CreateAttachmentResponse = "CreateAttachmentResponse";
        public const string CreateAttachmentResponseMessage = "CreateAttachmentResponseMessage";

        // DeleteAttachment
        public const string DeleteAttachment = "DeleteAttachment";
        public const string DeleteAttachmentResponse = "DeleteAttachmentResponse";
        public const string DeleteAttachmentResponseMessage = "DeleteAttachmentResponseMessage";

        // ResolveNames
        public const string ResolveNames = "ResolveNames";
        public const string ResolveNamesResponse = "ResolveNamesResponse";
        public const string ResolveNamesResponseMessage = "ResolveNamesResponseMessage";

        // ExpandDL
        public const string ExpandDL = "ExpandDL";
        public const string ExpandDLResponse = "ExpandDLResponse";
        public const string ExpandDLResponseMessage = "ExpandDLResponseMessage";

        // ExportItems

        public const string ExportItems = "ExportItems";
        public const string ExportItemsResponse = "ExportItemsResponse";
        public const string ExportItemsResponseMessage = "ExportItemsResponseMessage";
        
        // Upload Items

        public const string UploadItemsResponseMessage = "UploadItemsResponseMessage";
        public const string UploadItemsResponse = "UploadItemsResponse";
        public const string UploadItems = "UploadItems";

        // Subscribe
        public const string Subscribe = "Subscribe";
        public const string SubscribeResponse = "SubscribeResponse";
        public const string SubscribeResponseMessage = "SubscribeResponseMessage";
        public const string SubscriptionRequest = "SubscriptionRequest";

        // Unsubscribe
        public const string Unsubscribe = "Unsubscribe";
        public const string UnsubscribeResponse = "UnsubscribeResponse";
        public const string UnsubscribeResponseMessage = "UnsubscribeResponseMessage";

        // GetEvents
        public const string GetEvents = "GetEvents";
        public const string GetEventsResponse = "GetEventsResponse";
        public const string GetEventsResponseMessage = "GetEventsResponseMessage";

        // GetStreamingEvents
        public const string GetStreamingEvents = "GetStreamingEvents";
        public const string GetStreamingEventsResponse = "GetStreamingEventsResponse";
        public const string GetStreamingEventsResponseMessage = "GetStreamingEventsResponseMessage";
        public const string ConnectionStatus = "ConnectionStatus";
        public const string ErrorSubscriptionIds = "ErrorSubscriptionIds";
        public const string ConnectionTimeout = "ConnectionTimeout";
        public const string HeartbeatFrequency = "HeartbeatFrequency";

        // SyncFolderItems
        public const string SyncFolderItems = "SyncFolderItems";
        public const string SyncFolderItemsResponse = "SyncFolderItemsResponse";
        public const string SyncFolderItemsResponseMessage = "SyncFolderItemsResponseMessage";

        // SyncFolderHierarchy
        public const string SyncFolderHierarchy = "SyncFolderHierarchy";
        public const string SyncFolderHierarchyResponse = "SyncFolderHierarchyResponse";
        public const string SyncFolderHierarchyResponseMessage = "SyncFolderHierarchyResponseMessage";

        // GetUserOofSettings
        public const string GetUserOofSettingsRequest = "GetUserOofSettingsRequest";
        public const string GetUserOofSettingsResponse = "GetUserOofSettingsResponse";

        // SetUserOofSettings
        public const string SetUserOofSettingsRequest = "SetUserOofSettingsRequest";
        public const string SetUserOofSettingsResponse = "SetUserOofSettingsResponse";

        // GetUserAvailability
        public const string GetUserAvailabilityRequest = "GetUserAvailabilityRequest";
        public const string GetUserAvailabilityResponse = "GetUserAvailabilityResponse";
        public const string FreeBusyResponseArray = "FreeBusyResponseArray";
        public const string FreeBusyResponse = "FreeBusyResponse";
        public const string SuggestionsResponse = "SuggestionsResponse";

        // GetRoomLists
        public const string GetRoomListsRequest = "GetRoomLists";
        public const string GetRoomListsResponse = "GetRoomListsResponse";

        // GetRooms
        public const string GetRoomsRequest = "GetRooms";
        public const string GetRoomsResponse = "GetRoomsResponse";

        // ConvertId
        public const string ConvertId = "ConvertId";
        public const string ConvertIdResponse = "ConvertIdResponse";
        public const string ConvertIdResponseMessage = "ConvertIdResponseMessage";

        // AddDelegate
        public const string AddDelegate = "AddDelegate";
        public const string AddDelegateResponse = "AddDelegateResponse";
        public const string DelegateUserResponseMessageType = "DelegateUserResponseMessageType";

        // RemoveDelegte
        public const string RemoveDelegate = "RemoveDelegate";
        public const string RemoveDelegateResponse = "RemoveDelegateResponse";

        // GetDelegate
        public const string GetDelegate = "GetDelegate";
        public const string GetDelegateResponse = "GetDelegateResponse";

        // UpdateDelegate
        public const string UpdateDelegate = "UpdateDelegate";
        public const string UpdateDelegateResponse = "UpdateDelegateResponse";

        // CreateUserConfiguration
        public const string CreateUserConfiguration = "CreateUserConfiguration";
        public const string CreateUserConfigurationResponse = "CreateUserConfigurationResponse";
        public const string CreateUserConfigurationResponseMessage = "CreateUserConfigurationResponseMessage";

        // DeleteUserConfiguration
        public const string DeleteUserConfiguration = "DeleteUserConfiguration";
        public const string DeleteUserConfigurationResponse = "DeleteUserConfigurationResponse";
        public const string DeleteUserConfigurationResponseMessage = "DeleteUserConfigurationResponseMessage";

        // GetUserConfiguration
        public const string GetUserConfiguration = "GetUserConfiguration";
        public const string GetUserConfigurationResponse = "GetUserConfigurationResponse";
        public const string GetUserConfigurationResponseMessage = "GetUserConfigurationResponseMessage";

        // UpdateUserConfiguration
        public const string UpdateUserConfiguration = "UpdateUserConfiguration";
        public const string UpdateUserConfigurationResponse = "UpdateUserConfigurationResponse";
        public const string UpdateUserConfigurationResponseMessage = "UpdateUserConfigurationResponseMessage";

        // PlayOnPhone
        public const string PlayOnPhone = "PlayOnPhone";
        public const string PlayOnPhoneResponse = "PlayOnPhoneResponse";

        // GetPhoneCallInformation
        public const string GetPhoneCall = "GetPhoneCallInformation";
        public const string GetPhoneCallResponse = "GetPhoneCallInformationResponse";

        // DisconnectCall
        public const string DisconnectPhoneCall = "DisconnectPhoneCall";
        public const string DisconnectPhoneCallResponse = "DisconnectPhoneCallResponse";

        // GetServerTimeZones
        public const string GetServerTimeZones = "GetServerTimeZones";
        public const string GetServerTimeZonesResponse = "GetServerTimeZonesResponse";
        public const string GetServerTimeZonesResponseMessage = "GetServerTimeZonesResponseMessage";

        // GetInboxRules
        public const string GetInboxRules = "GetInboxRules";
        public const string GetInboxRulesResponse = "GetInboxRulesResponse";

        // UpdateInboxRules
        public const string UpdateInboxRules = "UpdateInboxRules";
        public const string UpdateInboxRulesResponse = "UpdateInboxRulesResponse";

        // ExecuteDiagnosticMethod
        public const string ExecuteDiagnosticMethod = "ExecuteDiagnosticMethod";
        public const string ExecuteDiagnosticMethodResponse = "ExecuteDiagnosticMethodResponse";
        public const string ExecuteDiagnosticMethodResponseMEssage = "ExecuteDiagnosticMethodResponseMessage";

        //GetPasswordExpirationDate
        public const string GetPasswordExpirationDateRequest = "GetPasswordExpirationDate";
        public const string GetPasswordExpirationDateResponse = "GetPasswordExpirationDateResponse";

        // GetSearchableMailboxes
        public const string GetSearchableMailboxes = "GetSearchableMailboxes";
        public const string GetSearchableMailboxesResponse = "GetSearchableMailboxesResponse";

        // GetDiscoverySearchConfiguration
        public const string GetDiscoverySearchConfiguration = "GetDiscoverySearchConfiguration";
        public const string GetDiscoverySearchConfigurationResponse = "GetDiscoverySearchConfigurationResponse";

        // GetHoldOnMailboxes
        public const string GetHoldOnMailboxes = "GetHoldOnMailboxes";
        public const string GetHoldOnMailboxesResponse = "GetHoldOnMailboxesResponse";

        // SetHoldOnMailboxes
        public const string SetHoldOnMailboxes = "SetHoldOnMailboxes";
        public const string SetHoldOnMailboxesResponse = "SetHoldOnMailboxesResponse";

        // SearchMailboxes
        public const string SearchMailboxes = "SearchMailboxes";
        public const string SearchMailboxesResponse = "SearchMailboxesResponse";
        public const string SearchMailboxesResponseMessage = "SearchMailboxesResponseMessage";

        // GetNonIndexableItemDetails
        public const string GetNonIndexableItemDetails = "GetNonIndexableItemDetails";
        public const string GetNonIndexableItemDetailsResponse = "GetNonIndexableItemDetailsResponse";

        // GetNonIndexableItemStatistics
        public const string GetNonIndexableItemStatistics = "GetNonIndexableItemStatistics";
        public const string GetNonIndexableItemStatisticsResponse = "GetNonIndexableItemStatisticsResponse";

        // eDiscovery
        public const string SearchQueries = "SearchQueries";
        public const string SearchQuery = "SearchQuery";
        public const string MailboxQuery = "MailboxQuery";
        public const string Query = "Query";
        public const string MailboxSearchScopes = "MailboxSearchScopes";
        public const string MailboxSearchScope = "MailboxSearchScope";
        public const string SearchScope = "SearchScope";
        public const string ResultType = "ResultType";
        public const string SortBy = "SortBy";
        public const string Order = "Order";
        public const string Language = "Language";
        public const string Deduplication = "Deduplication";
        public const string PageSize = "PageSize";
        public const string PageItemReference = "PageItemReference";
        public const string PageDirection = "PageDirection";
        public const string PreviewItemResponseShape = "PreviewItemResponseShape";
        public const string ExtendedProperties = "ExtendedProperties";
        public const string PageItemSize = "PageItemSize";
        public const string PageItemCount = "PageItemCount";
        public const string ItemCount = "ItemCount";
        public const string KeywordStats = "KeywordStats";
        public const string KeywordStat = "KeywordStat";
        public const string Keyword = "Keyword";
        public const string ItemHits = "ItemHits";
        public const string SearchPreviewItem = "SearchPreviewItem";
        public const string ChangeKey = "ChangeKey";
        public const string ParentId = "ParentId";
        public const string MailboxId = "MailboxId";
        public const string UniqueHash = "UniqueHash";
        public const string SortValue = "SortValue";
        public const string OwaLink = "OwaLink";
        public const string SmtpAddress = "SmtpAddress";
        public const string CreatedTime = "CreatedTime";
        public const string ReceivedTime = "ReceivedTime";
        public const string SentTime = "SentTime";
        public const string Preview = "Preview";
        public const string HasAttachment = "HasAttachment";
        public const string FailedMailboxes = "FailedMailboxes";
        public const string FailedMailbox = "FailedMailbox";
        public const string Token = "Token";
        public const string Refiners = "Refiners";
        public const string Refiner = "Refiner";
        public const string MailboxStats = "MailboxStats";
        public const string MailboxStat = "MailboxStat";
        public const string HoldId = "HoldId";
        public const string ActionType = "ActionType";
        public const string Mailboxes = "Mailboxes";
        public const string SearchFilter = "SearchFilter";
        public const string ReferenceId = "ReferenceId";
        public const string IsMembershipGroup = "IsMembershipGroup";
        public const string ExpandGroupMembership = "ExpandGroupMembership";
        public const string SearchableMailboxes = "SearchableMailboxes";
        public const string SearchableMailbox = "SearchableMailbox";
        public const string SearchMailboxesResult = "SearchMailboxesResult";
        public const string MailboxHoldResult = "MailboxHoldResult";
        public const string Statuses = "Statuses";
        public const string MailboxHoldStatuses = "MailboxHoldStatuses";
        public const string MailboxHoldStatus = "MailboxHoldStatus";
        public const string AdditionalInfo = "AdditionalInfo";
        public const string NonIndexableItemDetail = "NonIndexableItemDetail";
        public const string NonIndexableItemStatistic = "NonIndexableItemStatistic";
        public const string NonIndexableItemDetails = "NonIndexableItemDetails";
        public const string NonIndexableItemStatistics = "NonIndexableItemStatistics";
        public const string NonIndexableItemDetailsResult = "NonIndexableItemDetailsResult";
        public const string SearchArchiveOnly = "SearchArchiveOnly";
        public const string ErrorDescription = "ErrorDescription";
        public const string IsPartiallyIndexed = "IsPartiallyIndexed";
        public const string IsPermanentFailure = "IsPermanentFailure";
        public const string AttemptCount = "AttemptCount";
        public const string LastAttemptTime = "LastAttemptTime";
        public const string SearchId = "SearchId";
        public const string DiscoverySearchConfigurations = "DiscoverySearchConfigurations";
        public const string DiscoverySearchConfiguration = "DiscoverySearchConfiguration";
        public const string InPlaceHoldConfigurationOnly = "InPlaceHoldConfigurationOnly";
        public const string InPlaceHoldIdentity = "InPlaceHoldIdentity";
        public const string ItemHoldPeriod = "ItemHoldPeriod";
        public const string ManagedByOrganization = "ManagedByOrganization";
        public const string IsExternalMailbox = "IsExternalMailbox";
        public const string ExternalEmailAddress = "ExternalEmailAddress";
        public const string ExtendedAttributes = "ExtendedAttributes";
        public const string ExtendedAttribute = "ExtendedAttribute";
        public const string ExtendedAttributeName = "Name";
        public const string ExtendedAttributeValue = "Value";
        public const string SearchScopeType = "SearchScopeType";

        // GetAppManifests
        public const string GetAppManifestsRequest = "GetAppManifests";
        public const string GetAppManifestsResponse = "GetAppManifestsResponse";
        public const string Manifests = "Manifests";
        public const string Manifest = "Manifest";

        // GetAppManifests for TargetServerVersion > 2.5
        public const string Apps = "Apps";
        public const string App = "App";
        public const string Metadata = "Metadata";
        public const string ActionUrl = "ActionUrl";
        public const string AppStatus = "AppStatus";
        public const string EndNodeUrl = "EndNodeUrl";

        // GetClientExtension/SetClientExtension
        public const string GetClientExtensionRequest = "GetClientExtension";
        public const string ClientExtensionUserRequest = "UserParameters";
        public const string ClientExtensionUserEnabled = "UserEnabledExtensions";
        public const string ClientExtensionUserDisabled = "UserDisabledExtensions";
        public const string ClientExtensionRequestedIds = "RequestedExtensionIds";
        public const string ClientExtensionIsDebug = "IsDebug";
        public const string ClientExtensionRawMasterTableXml = "RawMasterTableXml";
        public const string GetClientExtensionResponse = "GetClientExtensionResponse";
        public const string ClientExtensionSpecificUsers = "SpecificUsers";
        public const string ClientExtensions = "ClientExtensions";
        public const string ClientExtension = "ClientExtension";
        public const string SetClientExtensionRequest = "SetClientExtension";
        public const string SetClientExtensionActions = "Actions";
        public const string SetClientExtensionAction = "Action";
        public const string SetClientExtensionResponse = "SetClientExtensionResponse";
        public const string SetClientExtensionResponseMessage = "SetClientExtensionResponseMessage";

        // GetOMEConfiguration/SetOMEConfiguration
        public const string GetOMEConfigurationRequest = "GetOMEConfiguration";
        public const string SetOMEConfigurationRequest = "SetOMEConfiguration";
        public const string OMEConfigurationXml = "Xml";
        public const string GetOMEConfigurationResponse = "GetOMEConfigurationResponse";
        public const string SetOMEConfigurationResponse = "SetOMEConfigurationResponse";

        // InstallApp
        public const string InstallAppRequest = "InstallApp";
        public const string InstallAppResponse = "InstallAppResponse";
        public const string MarketplaceAssetId = "MarketplaceAssetId";
        public const string MarketplaceContentMarket = "MarketplaceContentMarket";
        public const string SendWelcomeEmail = "SendWelcomeEmail";
        public const string WasFirstInstall = "WasFirstInstall";

        // UninstallApp
        public const string UninstallAppRequest = "UninstallApp";
        public const string UninstallAppResponse = "UninstallAppResponse";

        // DisableApp
        public const string DisableAppRequest = "DisableApp";
        public const string DisableAppResponse = "DisableAppResponse";

        // RegisterConsent
        public const string RegisterConsentRequest = "RegisterConsent";
        public const string RegisterConsentResponse = "RegisterConsentResponse";

        // GetAppMarketplaceUrl
        public const string GetAppMarketplaceUrlRequest = "GetAppMarketplaceUrl";
        public const string GetAppMarketplaceUrlResponse = "GetAppMarketplaceUrlResponse";

        // GetUserRetentionPolicyTags
        public const string GetUserRetentionPolicyTags = "GetUserRetentionPolicyTags";
        public const string GetUserRetentionPolicyTagsResponse = "GetUserRetentionPolicyTagsResponse";

        // MRM
        public const string RetentionPolicyTags = "RetentionPolicyTags";
        public const string RetentionPolicyTag = "RetentionPolicyTag";
        public const string RetentionId = "RetentionId";
        public const string RetentionPeriod = "RetentionPeriod";
        public const string RetentionAction = "RetentionAction";
        public const string Description = "Description";
        public const string IsVisible = "IsVisible";
        public const string OptedInto = "OptedInto";
        public const string IsArchive = "IsArchive";

        #endregion

        #region Groups

        // Like
        public const string Likers = "Likers";

        // GetUserUnifiedGroups
        public const string GetUserUnifiedGroups = "GetUserUnifiedGroups";
        public const string RequestedGroupsSets = "RequestedGroupsSets";
        public const string RequestedUnifiedGroupsSetItem = "RequestedUnifiedGroupsSet";
        public const string SortType = "SortType";
        public const string FilterType = "FilterType";
        public const string SortDirection = "SortDirection";
        public const string GroupsLimit = "GroupsLimit";
        public const string UserSmtpAddress = "UserSmtpAddress";

        public const string GetUserUnifiedGroupsResponseMessage = "GetUserUnifiedGroupsResponseMessage";
        public const string GroupsSets = "GroupsSets";
        public const string UnifiedGroupsSet = "UnifiedGroupsSet";
        public const string TotalGroups = "TotalGroups";
        public const string GroupsTag = "Groups";
        public const string UnifiedGroup = "UnifiedGroup";
        public const string MailboxGuid = "MailboxGuid";
        public const string LastVisitedTimeUtc = "LastVisitedTimeUtc";
        public const string AccessType = "AccessType";
        public const string ExternalDirectoryObjectId = "ExternalDirectoryObjectId";

        // GetUnifiedGroupUnseenCount
        public const string GetUnifiedGroupUnseenCount = "GetUnifiedGroupUnseenCount";
        public const string GroupIdentity = "GroupIdentity";
        public const string GroupIdentityType = "IdentityType";
        public const string GroupIdentityValue = "Value";

        public const string GetUnifiedGroupUnseenCountResponseMessage = "GetUnifiedGroupUnseenCountResponseMessage";
        public const string UnseenCount = "UnseenCount";

        // SetUnifiedGroupLastVisitedTimeRequest
        public const string SetUnifiedGroupLastVisitedTime = "SetUnifiedGroupLastVisitedTime";
        public const string SetUnifiedGroupLastVisitedTimeResponseMessage = "SetUnifiedGroupLastVisitedTimeResponseMessage";
        #endregion

        #region Hashtag and Mentions

        public const string Hashtags = "Hashtags";

        public const string Mentions = "Mentions";

        public const string MentionedMe = "MentionedMe";

        #endregion
        #region SOAP element names

        public const string SOAPEnvelopeElementName = "Envelope";
        public const string SOAPHeaderElementName = "Header";
        public const string SOAPBodyElementName = "Body";
        public const string SOAPFaultElementName = "Fault";
        public const string SOAPFaultCodeElementName = "faultcode";
        public const string SOAPFaultStringElementName = "faultstring";
        public const string SOAPFaultActorElementName = "faultactor";
        public const string SOAPDetailElementName = "detail";
        public const string EwsResponseCodeElementName = "ResponseCode";
        public const string EwsMessageElementName = "Message";
        public const string EwsLineElementName = "Line";
        public const string EwsPositionElementName = "Position";
        public const string EwsErrorCodeElementName = "ErrorCode";         // Generated by Availability
        public const string EwsExceptionTypeElementName = "ExceptionType"; // Generated by UM

        #endregion

    }
}