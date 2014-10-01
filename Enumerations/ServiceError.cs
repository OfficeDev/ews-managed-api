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
// <summary>Defines the ServiceError enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the error codes that can be returned by the Exchange Web Services.
    /// </summary>
    public enum ServiceError
    {
        /// <summary>
        /// NoError. Indicates that an error has not occurred.
        /// </summary>
        NoError = 0,

        /// <summary>
        /// Access is denied. Check credentials and try again.
        /// </summary>
        ErrorAccessDenied,

        /// <summary>
        /// The impersonation authentication header should not be included.
        /// </summary>
        ErrorAccessModeSpecified,

        /// <summary>
        /// Account is disabled. Contact the account administrator.
        /// </summary>
        ErrorAccountDisabled,

        /// <summary>
        /// Failed to add one or more delegates.
        /// </summary>
        ErrorAddDelegatesFailed,

        /// <summary>
        /// ErrorAddressSpaceNotFound
        /// </summary>
        ErrorAddressSpaceNotFound,

        /// <summary>
        /// Active Directory operation did not succeed. Try again later.
        /// </summary>
        ErrorADOperation,

        /// <summary>
        /// Invalid search criteria.
        /// </summary>
        ErrorADSessionFilter,

        /// <summary>
        /// Active Directory is unavailable. Try again later.
        /// </summary>
        ErrorADUnavailable,

        /// <summary>
        /// AffectedTaskOccurrences attribute is required for Task items.
        /// </summary>
        ErrorAffectedTaskOccurrencesRequired,

        /// <summary>
        /// The conversation action alwayscategorize or alwaysmove or alwaysdelete has failed.
        /// </summary>
        ErrorApplyConversationActionFailed,

        /// <summary>
        /// Archive mailbox not enabled
        /// </summary>
        ErrorArchiveMailboxNotEnabled,

        /// <summary>
        /// Unable to create the folder in archive mailbox to which the items will be archived
        /// </summary>
        ErrorArchiveFolderPathCreation,

        /// <summary>
        /// Unable to discover archive mailbox
        /// </summary>
        ErrorArchiveMailboxServiceDiscoveryFailed,

        /// <summary>
        /// The item has attachment at more than the maximum supported nest level.
        /// </summary>
        ErrorAttachmentNestLevelLimitExceeded,

        /// <summary>
        /// The file attachment exceeds the maximum supported size.
        /// </summary>
        ErrorAttachmentSizeLimitExceeded,

        /// <summary>
        /// ErrorAutoDiscoverFailed
        /// </summary>
        ErrorAutoDiscoverFailed,

        /// <summary>
        /// ErrorAvailabilityConfigNotFound
        /// </summary>
        ErrorAvailabilityConfigNotFound,

        /// <summary>
        /// Item was not processed as a result of a previous error.
        /// </summary>
        ErrorBatchProcessingStopped,

        /// <summary>
        /// Can not move or copy a calendar occurrence.
        /// </summary>
        ErrorCalendarCannotMoveOrCopyOccurrence,

        /// <summary>
        /// Cannot update calendar item that has already been deleted.
        /// </summary>
        ErrorCalendarCannotUpdateDeletedItem,

        /// <summary>
        /// The Id specified does not represent an occurrence.
        /// </summary>
        ErrorCalendarCannotUseIdForOccurrenceId,

        /// <summary>
        /// The specified Id does not represent a recurring master item.
        /// </summary>
        ErrorCalendarCannotUseIdForRecurringMasterId,

        /// <summary>
        /// Calendar item duration is too long.
        /// </summary>
        ErrorCalendarDurationIsTooLong,

        /// <summary>
        /// EndDate is earlier than StartDate
        /// </summary>
        ErrorCalendarEndDateIsEarlierThanStartDate,

        /// <summary>
        /// Cannot request CalendarView for the folder.
        /// </summary>
        ErrorCalendarFolderIsInvalidForCalendarView,

        /// <summary>
        /// Attribute has an invalid value.
        /// </summary>
        ErrorCalendarInvalidAttributeValue,

        /// <summary>
        /// The value of the DaysOfWeek property is not valid for time change pattern of time zone.
        /// </summary>
        ErrorCalendarInvalidDayForTimeChangePattern,

        /// <summary>
        /// The value of the DaysOfWeek property is invalid for a weekly recurrence.
        /// </summary>
        ErrorCalendarInvalidDayForWeeklyRecurrence,

        /// <summary>
        /// The property has invalid state.
        /// </summary>
        ErrorCalendarInvalidPropertyState,

        /// <summary>
        /// The property has an invalid value.
        /// </summary>
        ErrorCalendarInvalidPropertyValue,

        /// <summary>
        /// The recurrence is invalid.
        /// </summary>
        ErrorCalendarInvalidRecurrence,

        /// <summary>
        /// TimeZone is invalid.
        /// </summary>
        ErrorCalendarInvalidTimeZone,

        /// <summary>
        /// A meeting that's been canceled can't be accepted.
        /// </summary>
        ErrorCalendarIsCancelledForAccept,

        /// <summary>
        /// A canceled meeting can't be declined.
        /// </summary>
        ErrorCalendarIsCancelledForDecline,

        /// <summary>
        /// A canceled meeting can't be removed.
        /// </summary>
        ErrorCalendarIsCancelledForRemove,

        /// <summary>
        /// A canceled meeting can't be accepted tentatively.
        /// </summary>
        ErrorCalendarIsCancelledForTentative,

        /// <summary>
        /// AcceptItem action is invalid for a delegated meeting message.
        /// </summary>
        ErrorCalendarIsDelegatedForAccept,

        /// <summary>
        /// DeclineItem operation is invalid for a delegated meeting message.
        /// </summary>
        ErrorCalendarIsDelegatedForDecline,

        /// <summary>
        /// RemoveItem action is invalid for a delegated meeting message.
        /// </summary>
        ErrorCalendarIsDelegatedForRemove,

        /// <summary>
        /// The TentativelyAcceptItem action isn't valid for a delegated meeting message.
        /// </summary>
        ErrorCalendarIsDelegatedForTentative,

        /// <summary>
        /// User must be an organizer for CancelCalendarItem action.
        /// </summary>
        ErrorCalendarIsNotOrganizer,

        /// <summary>
        /// The user is the organizer of this meeting, and cannot, therefore, accept it.
        /// </summary>
        ErrorCalendarIsOrganizerForAccept,

        /// <summary>
        /// The user is the organizer of this meeting, and cannot, therefore, decline it.
        /// </summary>
        ErrorCalendarIsOrganizerForDecline,

        /// <summary>
        /// The user is the organizer of this meeting, and cannot, therefore, remove it.
        /// </summary>
        ErrorCalendarIsOrganizerForRemove,

        /// <summary>
        /// The user is the organizer of this meeting, and therefore can't tentatively accept it.
        /// </summary>
        ErrorCalendarIsOrganizerForTentative,

        /// <summary>
        /// The meeting request is out of date. The calendar couldn't be updated.
        /// </summary>
        ErrorCalendarMeetingRequestIsOutOfDate,

        /// <summary>
        /// Occurrence index is out of recurrence range.
        /// </summary>
        ErrorCalendarOccurrenceIndexIsOutOfRecurrenceRange,

        /// <summary>
        /// Occurrence with this index was previously deleted from the recurrence.
        /// </summary>
        ErrorCalendarOccurrenceIsDeletedFromRecurrence,

        /// <summary>
        /// The calendar property falls out of valid range.
        /// </summary>
        ErrorCalendarOutOfRange,

        /// <summary>
        /// The specified view range exceeds the maximum range of two years.
        /// </summary>
        ErrorCalendarViewRangeTooBig,

        /// <summary>
        /// Failed to get valid Active Directory information for the calling account. Confirm that it
        /// is a valid Active Directory account.
        /// </summary>
        ErrorCallerIsInvalidADAccount,

        /// <summary>
        /// Cannot archive items in Calendar, contact to task folders
        /// </summary>
        ErrorCannotArchiveCalendarContactTaskFolderException,

        /// <summary>
        /// Cannot archive items in archive mailboxes
        /// </summary>
        ErrorCannotArchiveItemsInArchiveMailbox,

        /// <summary>
        /// Cannot archive items in public folders
        /// </summary>
        ErrorCannotArchiveItemsInPublicFolders,

        /// <summary>
        /// Cannot create a calendar item in a non-calendar folder.
        /// </summary>
        ErrorCannotCreateCalendarItemInNonCalendarFolder,

        /// <summary>
        /// Cannot create a contact in a non-contact folder.
        /// </summary>
        ErrorCannotCreateContactInNonContactFolder,

        /// <summary>
        /// Cannot create a post item in a folder that is not a mail folder.
        /// </summary>
        ErrorCannotCreatePostItemInNonMailFolder,

        /// <summary>
        /// Cannot create a task in a non-task Folder.
        /// </summary>
        ErrorCannotCreateTaskInNonTaskFolder,

        /// <summary>
        /// Object cannot be deleted.
        /// </summary>
        ErrorCannotDeleteObject,

        /// <summary>
        /// Deleting a task occurrence is not permitted on non-recurring tasks, on the last
        /// occurrence of a recurring task or on a regenerating task.
        /// </summary>
        ErrorCannotDeleteTaskOccurrence,

        /// <summary>
        /// Mandatory extensions cannot be disabled by end users
        /// </summary>
        ErrorCannotDisableMandatoryExtension,

        /// <summary>
        /// Folder cannot be emptied.
        /// </summary>
        ErrorCannotEmptyFolder,

        /// <summary>
        /// Cannot get external ECP URL. This might happen if external ECP URL isn't configured
        /// </summary>
        ErrorCannotGetExternalEcpUrl,

        /// <summary>
        /// Unable to read the folder path for the source folder while archiving items
        /// </summary>
        ErrorCannotGetSourceFolderPath,

        /// <summary>
        /// The attachment could not be opened.
        /// </summary>
        ErrorCannotOpenFileAttachment,

        /// <summary>
        /// Expected a PermissionSet but received a CalendarPermissionSet.
        /// </summary>
        ErrorCannotSetCalendarPermissionOnNonCalendarFolder,

        /// <summary>
        /// Expected a CalendarPermissionSet but received a PermissionSet.
        /// </summary>
        ErrorCannotSetNonCalendarPermissionOnCalendarFolder,

        /// <summary>
        /// Cannot set UnknownEntries on a PermissionSet or CalendarPermissionSet.
        /// </summary>
        ErrorCannotSetPermissionUnknownEntries,

        /// <summary>
        /// Cannot specify search folders as source folders while archiving items
        /// </summary>
        ErrorCannotSpecifySearchFolderAsSourceFolder,

        /// <summary>
        /// Expected an item Id but received a folder Id.
        /// </summary>
        ErrorCannotUseFolderIdForItemId,

        /// <summary>
        /// Expected a folder Id but received an item Id.
        /// </summary>
        ErrorCannotUseItemIdForFolderId,

        /// <summary>
        /// ChangeKey is required if overriding automatic conflict resolution.
        /// </summary>
        ErrorChangeKeyRequired,

        /// <summary>
        /// ChangeKey is required for this operation.
        /// </summary>
        ErrorChangeKeyRequiredForWriteOperations,

        /// <summary>
        /// ErrorClientDisconnected
        /// </summary>
        ErrorClientDisconnected,

        /// <summary>
        /// Connection did not succeed. Try again later.
        /// </summary>
        ErrorConnectionFailed,

        /// <summary>
        /// The Contains filter can only be used for string properties.
        /// </summary>
        ErrorContainsFilterWrongType,

        /// <summary>
        /// Content conversion failed.
        /// </summary>
        ErrorContentConversionFailed,

        /// <summary>
        /// Data is corrupt.
        /// </summary>
        ErrorCorruptData,

        /// <summary>
        /// Unable to create item. The user account does not have the right to create items.
        /// </summary>
        ErrorCreateItemAccessDenied,

        /// <summary>
        /// Failed to create one or more of the specified managed folders.
        /// </summary>
        ErrorCreateManagedFolderPartialCompletion,

        /// <summary>
        /// Unable to create subfolder. The user account does not have the right to create
        /// subfolders.
        /// </summary>
        ErrorCreateSubfolderAccessDenied,

        /// <summary>
        /// Move and Copy operations across mailbox boundaries are not permitted.
        /// </summary>
        ErrorCrossMailboxMoveCopy,

        /// <summary>
        /// This request isn't allowed because the Client Access server that's servicing the request
        /// is in a different site than the requested resource. Use Autodiscover to find the correct
        /// URL for accessing the specified resource.
        /// </summary>
        ErrorCrossSiteRequest,

        /// <summary>
        /// Property exceeds the maximum supported size.
        /// </summary>
        ErrorDataSizeLimitExceeded,

        /// <summary>
        /// Invalid data source operation.
        /// </summary>
        ErrorDataSourceOperation,

        /// <summary>
        /// The user is already a delegate for the mailbox.
        /// </summary>
        ErrorDelegateAlreadyExists,

        /// <summary>
        /// This is an invalid operation. Cannot add owner as delegate.
        /// </summary>
        ErrorDelegateCannotAddOwner,

        /// <summary>
        /// Delegate is not configured properly.
        /// </summary>
        ErrorDelegateMissingConfiguration,

        /// <summary>
        /// The delegate does not map to a user in the Active Directory.
        /// </summary>
        ErrorDelegateNoUser,

        /// <summary>
        /// Cannot add the delegate user. Failed to validate the changes.
        /// </summary>
        ErrorDelegateValidationFailed,

        /// <summary>
        /// Distinguished folders cannot be deleted.
        /// </summary>
        ErrorDeleteDistinguishedFolder,

        /// <summary>
        /// The deletion failed.
        /// </summary>
        ErrorDeleteItemsFailed,

        /// <summary>
        /// DistinguishedUser should not be specified for a Delegate User.
        /// </summary>
        ErrorDistinguishedUserNotSupported,

        /// <summary>
        /// The group member doesn't exist.
        /// </summary>
        ErrorDistributionListMemberNotExist,

        /// <summary>
        /// The specified list of managed folder names contains duplicate entries.
        /// </summary>
        ErrorDuplicateInputFolderNames,

        /// <summary>
        /// A duplicate exchange legacy DN.
        /// </summary>
        ErrorDuplicateLegacyDistinguishedName,

        /// <summary>
        /// A duplicate SOAP header was received.
        /// </summary>
        ErrorDuplicateSOAPHeader,

        /// <summary>
        /// The specified permission set contains duplicate UserIds.
        /// </summary>
        ErrorDuplicateUserIdsSpecified,

        /// <summary>
        /// The email address associated with a folder Id does not match the mailbox you are
        /// operating on.
        /// </summary>
        ErrorEmailAddressMismatch,

        /// <summary>
        /// The watermark used for creating this subscription was not found.
        /// </summary>
        ErrorEventNotFound,

        /// <summary>
        /// You have exceeded the available concurrent connections for your account.  Try again once
        /// your other requests have completed.
        /// </summary>
        ErrorExceededConnectionCount,

        /// <summary>
        /// You have exceeded the maximum number of objects that can be returned for the find
        /// operation. Use paging to reduce the result size and try your request again.
        /// </summary>
        ErrorExceededFindCountLimit,

        /// <summary>
        /// You have exceeded the available subscriptions for your account.  Remove unnecessary
        /// subscriptions and try your request again.
        /// </summary>
        ErrorExceededSubscriptionCount,

        /// <summary>
        /// Subscription information is not available. Subscription is expired.
        /// </summary>
        ErrorExpiredSubscription,

        /// <summary>
        /// Extension with id specified was not found
        /// </summary>
        ErrorExtensionNotFound,

        /// <summary>
        /// The folder is corrupt.
        /// </summary>
        ErrorFolderCorrupt,

        /// <summary>
        /// A folder with the specified name already exists.
        /// </summary>
        ErrorFolderExists,

        /// <summary>
        /// The specified folder could not be found in the store.
        /// </summary>
        ErrorFolderNotFound,

        /// <summary>
        /// ErrorFolderPropertRequestFailed
        /// </summary>
        ErrorFolderPropertRequestFailed,

        /// <summary>
        /// The folder save operation did not succeed.
        /// </summary>
        ErrorFolderSave,

        /// <summary>
        /// The save operation failed or partially succeeded.
        /// </summary>
        ErrorFolderSaveFailed,

        /// <summary>
        /// The folder save operation failed due to invalid property values.
        /// </summary>
        ErrorFolderSavePropertyError,

        /// <summary>
        /// ErrorFreeBusyDLLimitReached
        /// </summary>
        ErrorFreeBusyDLLimitReached,

        /// <summary>
        /// ErrorFreeBusyGenerationFailed
        /// </summary>
        ErrorFreeBusyGenerationFailed,

        /// <summary>
        /// ErrorGetServerSecurityDescriptorFailed
        /// </summary>
        ErrorGetServerSecurityDescriptorFailed,

        /// <summary>
        /// ErrorImContactLimitReached
        /// </summary>
        ErrorImContactLimitReached,

        /// <summary>
        /// ErrorImGroupDisplayNameAlreadyExists
        /// </summary>
        ErrorImGroupDisplayNameAlreadyExists,

        /// <summary>
        /// ErrorImGroupLimitReached
        /// </summary>
        ErrorImGroupLimitReached,

        /// <summary>
        /// The account does not have permission to impersonate the requested user.
        /// </summary>
        ErrorImpersonateUserDenied,

        /// <summary>
        /// ErrorImpersonationDenied
        /// </summary>
        ErrorImpersonationDenied,

        /// <summary>
        /// Impersonation failed.
        /// </summary>
        ErrorImpersonationFailed,

        /// <summary>
        /// ErrorInboxRulesValidationError
        /// </summary>
        ErrorInboxRulesValidationError,

        /// <summary>
        /// The request is valid but does not specify the correct server version in the
        /// RequestServerVersion SOAP header.  Ensure that the RequestServerVersion SOAP header is
        /// set with the correct RequestServerVersionValue.
        /// </summary>
        ErrorIncorrectSchemaVersion,

        /// <summary>
        /// An object within a change description must contain one and only one property to modify.
        /// </summary>
        ErrorIncorrectUpdatePropertyCount,

        /// <summary>
        /// ErrorIndividualMailboxLimitReached
        /// </summary>
        ErrorIndividualMailboxLimitReached,

        /// <summary>
        /// Resources are unavailable. Try again later.
        /// </summary>
        ErrorInsufficientResources,

        /// <summary>
        /// An internal server error occurred. The operation failed.
        /// </summary>
        ErrorInternalServerError,

        /// <summary>
        /// An internal server error occurred. Try again later.
        /// </summary>
        ErrorInternalServerTransientError,

        /// <summary>
        /// ErrorInvalidAccessLevel
        /// </summary>
        ErrorInvalidAccessLevel,

        /// <summary>
        /// ErrorInvalidArgument
        /// </summary>
        ErrorInvalidArgument,

        /// <summary>
        /// The specified attachment Id is invalid.
        /// </summary>
        ErrorInvalidAttachmentId,

        /// <summary>
        /// Attachment subfilters must have a single TextFilter therein.
        /// </summary>
        ErrorInvalidAttachmentSubfilter,

        /// <summary>
        /// Attachment subfilters must have a single TextFilter on the display name only.
        /// </summary>
        ErrorInvalidAttachmentSubfilterTextFilter,

        /// <summary>
        /// ErrorInvalidAuthorizationContext
        /// </summary>
        ErrorInvalidAuthorizationContext,

        /// <summary>
        /// The change key is invalid.
        /// </summary>
        ErrorInvalidChangeKey,

        /// <summary>
        /// ErrorInvalidClientSecurityContext
        /// </summary>
        ErrorInvalidClientSecurityContext,

        /// <summary>
        /// CompleteDate cannot be set to a date in the future.
        /// </summary>
        ErrorInvalidCompleteDate,

        /// <summary>
        /// The e-mail address that was supplied isn't valid.
        /// </summary>
        ErrorInvalidContactEmailAddress,

        /// <summary>
        /// The e-mail index supplied isn't valid.
        /// </summary>
        ErrorInvalidContactEmailIndex,

        /// <summary>
        /// ErrorInvalidCrossForestCredentials
        /// </summary>
        ErrorInvalidCrossForestCredentials,

        /// <summary>
        /// Invalid Delegate Folder Permission.
        /// </summary>
        ErrorInvalidDelegatePermission,

        /// <summary>
        /// One or more UserId parameters are invalid. Make sure that the PrimarySmtpAddress, Sid and
        /// DisplayName properties refer to the same user when specified.
        /// </summary>
        ErrorInvalidDelegateUserId,

        /// <summary>
        /// An ExchangeImpersonation SOAP header must contain a user principal name, user SID, or
        /// primary SMTP address.
        /// </summary>
        ErrorInvalidExchangeImpersonationHeaderData,

        /// <summary>
        /// Second operand in Excludes expression must be uint compatible.
        /// </summary>
        ErrorInvalidExcludesRestriction,

        /// <summary>
        /// FieldURI can only be used in Contains expressions.
        /// </summary>
        ErrorInvalidExpressionTypeForSubFilter,

        /// <summary>
        /// The extended property attribute combination is invalid.
        /// </summary>
        ErrorInvalidExtendedProperty,

        /// <summary>
        /// The extended property value is inconsistent with its type.
        /// </summary>
        ErrorInvalidExtendedPropertyValue,

        /// <summary>
        /// The original sender of the message (initiator field in the sharing metadata) is not
        /// valid.
        /// </summary>
        ErrorInvalidExternalSharingInitiator,

        /// <summary>
        /// The sharing message is not intended for this caller.
        /// </summary>
        ErrorInvalidExternalSharingSubscriber,

        /// <summary>
        /// The organization is either not federated, or it's configured incorrectly.
        /// </summary>
        ErrorInvalidFederatedOrganizationId,

        /// <summary>
        /// Folder Id is invalid.
        /// </summary>
        ErrorInvalidFolderId,

        /// <summary>
        /// ErrorInvalidFolderTypeForOperation
        /// </summary>
        ErrorInvalidFolderTypeForOperation,

        /// <summary>
        /// Invalid fractional paging offset values.
        /// </summary>
        ErrorInvalidFractionalPagingParameters,

        /// <summary>
        /// ErrorInvalidFreeBusyViewType
        /// </summary>
        ErrorInvalidFreeBusyViewType,

        /// <summary>
        /// Either DataType or SharedFolderId must be specified, but not both.
        /// </summary>
        ErrorInvalidGetSharingFolderRequest,

        /// <summary>
        /// The Id is invalid.
        /// </summary>
        ErrorInvalidId,

        /// <summary>
        /// The Im Contact id was invalid.
        /// </summary>
        ErrorInvalidImContactId,

        /// <summary>
        /// The Im Distribution Group Smtp Address was invalid.
        /// </summary>
        ErrorInvalidImDistributionGroupSmtpAddress,

        /// <summary>
        /// The Im Contact id was invalid.
        /// </summary>
        ErrorInvalidImGroupId,

        /// <summary>
        /// Id must be non-empty.
        /// </summary>
        ErrorInvalidIdEmpty,

        /// <summary>
        /// Id is malformed.
        /// </summary>
        ErrorInvalidIdMalformed,

        /// <summary>
        /// The EWS Id is in EwsLegacyId format which is not supported by the Exchange version
        /// specified by your request. Please use the ConvertId method to convert from EwsLegacyId 
        /// to EwsId format.
        /// </summary>
        ErrorInvalidIdMalformedEwsLegacyIdFormat,

        /// <summary>
        /// Moniker exceeded allowable length.
        /// </summary>
        ErrorInvalidIdMonikerTooLong,

        /// <summary>
        /// The Id does not represent an item attachment.
        /// </summary>
        ErrorInvalidIdNotAnItemAttachmentId,

        /// <summary>
        /// ResolveNames returned an invalid Id.
        /// </summary>
        ErrorInvalidIdReturnedByResolveNames,

        /// <summary>
        /// Id exceeded allowable length.
        /// </summary>
        ErrorInvalidIdStoreObjectIdTooLong,

        /// <summary>
        /// Too many attachment levels.
        /// </summary>
        ErrorInvalidIdTooManyAttachmentLevels,

        /// <summary>
        /// The Id Xml is invalid.
        /// </summary>
        ErrorInvalidIdXml,

        /// <summary>
        /// The specified indexed paging values are invalid.
        /// </summary>
        ErrorInvalidIndexedPagingParameters,

        /// <summary>
        /// Only one child node is allowed when setting an Internet Message Header.
        /// </summary>
        ErrorInvalidInternetHeaderChildNodes,

        /// <summary>
        /// Item type is invalid for AcceptItem action.
        /// </summary>
        ErrorInvalidItemForOperationAcceptItem,

        /// <summary>
        /// Item type is invalid for ArchiveItem action.
        /// </summary>
        ErrorInvalidItemForOperationArchiveItem,

        /// <summary>
        /// Item type is invalid for CancelCalendarItem action.
        /// </summary>
        ErrorInvalidItemForOperationCancelItem,

        /// <summary>
        /// Item type is invalid for CreateItem operation.
        /// </summary>
        ErrorInvalidItemForOperationCreateItem,

        /// <summary>
        /// Item type is invalid for CreateItemAttachment operation.
        /// </summary>
        ErrorInvalidItemForOperationCreateItemAttachment,

        /// <summary>
        /// Item type is invalid for DeclineItem operation.
        /// </summary>
        ErrorInvalidItemForOperationDeclineItem,

        /// <summary>
        /// ExpandDL operation does not support this item type.
        /// </summary>
        ErrorInvalidItemForOperationExpandDL,

        /// <summary>
        /// Item type is invalid for RemoveItem operation.
        /// </summary>
        ErrorInvalidItemForOperationRemoveItem,

        /// <summary>
        /// Item type is invalid for SendItem operation.
        /// </summary>
        ErrorInvalidItemForOperationSendItem,

        /// <summary>
        /// The item of this type is invalid for TentativelyAcceptItem action.
        /// </summary>
        ErrorInvalidItemForOperationTentative,

        /// <summary>
        /// The logon type isn't valid.
        /// </summary>
        ErrorInvalidLogonType,

        /// <summary>
        /// Mailbox is invalid. Verify the specified Mailbox property.
        /// </summary>
        ErrorInvalidMailbox,

        /// <summary>
        /// The Managed Folder property is corrupt or otherwise invalid.
        /// </summary>
        ErrorInvalidManagedFolderProperty,

        /// <summary>
        /// The managed folder has an invalid quota.
        /// </summary>
        ErrorInvalidManagedFolderQuota,

        /// <summary>
        /// The managed folder has an invalid storage limit value.
        /// </summary>
        ErrorInvalidManagedFolderSize,

        /// <summary>
        /// ErrorInvalidMergedFreeBusyInterval
        /// </summary>
        ErrorInvalidMergedFreeBusyInterval,

        /// <summary>
        /// The specified value is not a valid name for name resolution.
        /// </summary>
        ErrorInvalidNameForNameResolution,

        /// <summary>
        /// ErrorInvalidNetworkServiceContext
        /// </summary>
        ErrorInvalidNetworkServiceContext,

        /// <summary>
        /// ErrorInvalidOofParameter
        /// </summary>
        ErrorInvalidOofParameter,

        /// <summary>
        /// ErrorInvalidOperation
        /// </summary>
        ErrorInvalidOperation,

        /// <summary>
        /// ErrorInvalidOrganizationRelationshipForFreeBusy
        /// </summary>
        ErrorInvalidOrganizationRelationshipForFreeBusy,

        /// <summary>
        /// MaxEntriesReturned must be greater than zero.
        /// </summary>
        ErrorInvalidPagingMaxRows,

        /// <summary>
        /// Cannot create a subfolder within a SearchFolder.
        /// </summary>
        ErrorInvalidParentFolder,

        /// <summary>
        /// PercentComplete must be an integer between 0 and 100.
        /// </summary>
        ErrorInvalidPercentCompleteValue,

        /// <summary>
        /// The permission settings were not valid.
        /// </summary>
        ErrorInvalidPermissionSettings,

        /// <summary>
        /// The phone call ID isn't valid.
        /// </summary>
        ErrorInvalidPhoneCallId,

        /// <summary>
        /// The phone number isn't valid.
        /// </summary>
        ErrorInvalidPhoneNumber,

        /// <summary>
        /// The append action is not supported for this property.
        /// </summary>
        ErrorInvalidPropertyAppend,

        /// <summary>
        /// The delete action is not supported for this property.
        /// </summary>
        ErrorInvalidPropertyDelete,

        /// <summary>
        /// Property cannot be used in Exists expression.  Use IsEqualTo instead.
        /// </summary>
        ErrorInvalidPropertyForExists,

        /// <summary>
        /// Property is not valid for this operation.
        /// </summary>
        ErrorInvalidPropertyForOperation,

        /// <summary>
        /// Property is not valid for this object type.
        /// </summary>
        ErrorInvalidPropertyRequest,

        /// <summary>
        /// Set action is invalid for property.
        /// </summary>
        ErrorInvalidPropertySet,

        /// <summary>
        /// Update operation is invalid for property of a sent message.
        /// </summary>
        ErrorInvalidPropertyUpdateSentMessage,

        /// <summary>
        /// The proxy security context is invalid.
        /// </summary>
        ErrorInvalidProxySecurityContext,

        /// <summary>
        /// SubscriptionId is invalid. Subscription is not a pull subscription.
        /// </summary>
        ErrorInvalidPullSubscriptionId,

        /// <summary>
        /// URL specified for push subscription is invalid.
        /// </summary>
        ErrorInvalidPushSubscriptionUrl,

        /// <summary>
        /// One or more recipients are invalid.
        /// </summary>
        ErrorInvalidRecipients,

        /// <summary>
        /// Recipient subfilters are only supported when there are two expressions within a single
        /// AND filter.
        /// </summary>
        ErrorInvalidRecipientSubfilter,

        /// <summary>
        /// Recipient subfilter must have a comparison filter that tests equality to recipient type
        /// or attendee type.
        /// </summary>
        ErrorInvalidRecipientSubfilterComparison,

        /// <summary>
        /// Recipient subfilters must have a text filter and a comparison filter in that order.
        /// </summary>
        ErrorInvalidRecipientSubfilterOrder,

        /// <summary>
        /// Recipient subfilter must have a TextFilter on the SMTP address only.
        /// </summary>
        ErrorInvalidRecipientSubfilterTextFilter,

        /// <summary>
        /// The reference item does not support the requested operation.
        /// </summary>
        ErrorInvalidReferenceItem,

        /// <summary>
        /// The request is invalid.
        /// </summary>
        ErrorInvalidRequest,

        /// <summary>
        /// The restriction is invalid.
        /// </summary>
        ErrorInvalidRestriction,

        /// <summary>
        /// ErrorInvalidRetentionIdTagTypeMismatch.
        /// </summary>
        ErrorInvalidRetentionTagTypeMismatch,

        /// <summary>
        /// ErrorInvalidRetentionTagInvisible.
        /// </summary>
        ErrorInvalidRetentionTagInvisible,

        /// <summary>
        /// ErrorInvalidRetentionTagInheritance.
        /// </summary>
        ErrorInvalidRetentionTagInheritance,

        /// <summary>
        /// ErrorInvalidRetentionTagIdGuid.
        /// </summary>
        ErrorInvalidRetentionTagIdGuid,

        /// <summary>
        /// The routing type format is invalid.
        /// </summary>
        ErrorInvalidRoutingType,

        /// <summary>
        /// ErrorInvalidScheduledOofDuration
        /// </summary>
        ErrorInvalidScheduledOofDuration,

        /// <summary>
        /// The mailbox that was requested doesn't support the specified RequestServerVersion.
        /// </summary>
        ErrorInvalidSchemaVersionForMailboxVersion,

        /// <summary>
        /// ErrorInvalidSecurityDescriptor
        /// </summary>
        ErrorInvalidSecurityDescriptor,

        /// <summary>
        /// Invalid combination of SaveItemToFolder attribute and SavedItemFolderId element.
        /// </summary>
        ErrorInvalidSendItemSaveSettings,

        /// <summary>
        /// Invalid serialized access token.
        /// </summary>
        ErrorInvalidSerializedAccessToken,

        /// <summary>
        /// The specified server version is invalid.
        /// </summary>
        ErrorInvalidServerVersion,

        /// <summary>
        /// The sharing message metadata is not valid.
        /// </summary>
        ErrorInvalidSharingData,

        /// <summary>
        /// The sharing message is not valid.
        /// </summary>
        ErrorInvalidSharingMessage,

        /// <summary>
        /// A SID with an invalid format was encountered.
        /// </summary>
        ErrorInvalidSid,

        /// <summary>
        /// The SIP address isn't valid.
        /// </summary>
        ErrorInvalidSIPUri,

        /// <summary>
        /// The SMTP address format is invalid.
        /// </summary>
        ErrorInvalidSmtpAddress,

        /// <summary>
        /// Invalid subFilterType.
        /// </summary>
        ErrorInvalidSubfilterType,

        /// <summary>
        /// SubFilterType is not attendee type.
        /// </summary>
        ErrorInvalidSubfilterTypeNotAttendeeType,

        /// <summary>
        /// SubFilterType is not recipient type.
        /// </summary>
        ErrorInvalidSubfilterTypeNotRecipientType,

        /// <summary>
        /// Subscription is invalid.
        /// </summary>
        ErrorInvalidSubscription,

        /// <summary>
        /// A subscription can only be established on a single public folder or on folders from a
        /// single mailbox.
        /// </summary>
        ErrorInvalidSubscriptionRequest,

        /// <summary>
        /// Synchronization state data is corrupt or otherwise invalid.
        /// </summary>
        ErrorInvalidSyncStateData,

        /// <summary>
        /// ErrorInvalidTimeInterval
        /// </summary>
        ErrorInvalidTimeInterval,

        /// <summary>
        /// A UserId was not valid.
        /// </summary>
        ErrorInvalidUserInfo,

        /// <summary>
        /// ErrorInvalidUserOofSettings
        /// </summary>
        ErrorInvalidUserOofSettings,

        /// <summary>
        /// The impersonation principal name is invalid.
        /// </summary>
        ErrorInvalidUserPrincipalName,

        /// <summary>
        /// The user SID is invalid or does not map to a user in the Active Directory.
        /// </summary>
        ErrorInvalidUserSid,

        /// <summary>
        /// ErrorInvalidUserSidMissingUPN
        /// </summary>
        ErrorInvalidUserSidMissingUPN,

        /// <summary>
        /// The specified value is invalid for property.
        /// </summary>
        ErrorInvalidValueForProperty,

        /// <summary>
        /// The watermark is invalid.
        /// </summary>
        ErrorInvalidWatermark,

        /// <summary>
        /// A valid IP gateway couldn't be found.
        /// </summary>
        ErrorIPGatewayNotFound,

        /// <summary>
        /// The send or update operation could not be performed because the change key passed in the
        /// request does not match the current change key for the item.
        /// </summary>
        ErrorIrresolvableConflict,

        /// <summary>
        /// The item is corrupt.
        /// </summary>
        ErrorItemCorrupt,

        /// <summary>
        /// The specified object was not found in the store.
        /// </summary>
        ErrorItemNotFound,

        /// <summary>
        /// One or more of the properties requested for this item could not be retrieved.
        /// </summary>
        ErrorItemPropertyRequestFailed,

        /// <summary>
        /// The item save operation did not succeed.
        /// </summary>
        ErrorItemSave,

        /// <summary>
        /// Item save operation did not succeed.
        /// </summary>
        ErrorItemSavePropertyError,

        /// <summary>
        /// ErrorLegacyMailboxFreeBusyViewTypeNotMerged
        /// </summary>
        ErrorLegacyMailboxFreeBusyViewTypeNotMerged,

        /// <summary>
        /// ErrorLocalServerObjectNotFound
        /// </summary>
        ErrorLocalServerObjectNotFound,

        /// <summary>
        /// ErrorLogonAsNetworkServiceFailed
        /// </summary>
        ErrorLogonAsNetworkServiceFailed,

        /// <summary>
        /// Unable to access an account or mailbox.
        /// </summary>
        ErrorMailboxConfiguration,

        /// <summary>
        /// ErrorMailboxDataArrayEmpty
        /// </summary>
        ErrorMailboxDataArrayEmpty,

        /// <summary>
        /// ErrorMailboxDataArrayTooBig
        /// </summary>
        ErrorMailboxDataArrayTooBig,

        /// <summary>
        /// ErrorMailboxFailover
        /// </summary>
        ErrorMailboxFailover,

        /// <summary>
        /// The specific mailbox hold is not found.
        /// </summary>
        ErrorMailboxHoldNotFound,

        /// <summary>
        /// ErrorMailboxLogonFailed
        /// </summary>
        ErrorMailboxLogonFailed,

        /// <summary>
        /// Mailbox move in progress. Try again later.
        /// </summary>
        ErrorMailboxMoveInProgress,

        /// <summary>
        /// The mailbox database is temporarily unavailable.
        /// </summary>
        ErrorMailboxStoreUnavailable,

        /// <summary>
        /// ErrorMailRecipientNotFound
        /// </summary>
        ErrorMailRecipientNotFound,

        /// <summary>
        /// MailTips aren't available for your organization.
        /// </summary>
        ErrorMailTipsDisabled,

        /// <summary>
        /// The specified Managed Folder already exists in the mailbox.
        /// </summary>
        ErrorManagedFolderAlreadyExists,

        /// <summary>
        /// Unable to find the specified managed folder in the Active Directory.
        /// </summary>
        ErrorManagedFolderNotFound,

        /// <summary>
        /// Failed to create or bind to the folder: Managed Folders
        /// </summary>
        ErrorManagedFoldersRootFailure,

        /// <summary>
        /// ErrorMeetingSuggestionGenerationFailed
        /// </summary>
        ErrorMeetingSuggestionGenerationFailed,

        /// <summary>
        /// MessageDisposition attribute is required.
        /// </summary>
        ErrorMessageDispositionRequired,

        /// <summary>
        /// The message exceeds the maximum supported size.
        /// </summary>
        ErrorMessageSizeExceeded,

        /// <summary>
        /// The domain specified in the tracking request doesn't exist.
        /// </summary>
        ErrorMessageTrackingNoSuchDomain,

        /// <summary>
        /// The log search service can't track this message.
        /// </summary>
        ErrorMessageTrackingPermanentError,

        /// <summary>
        /// The log search service isn't currently available. Please try again later.
        /// </summary>
        ErrorMessageTrackingTransientError,

        /// <summary>
        /// MIME content conversion failed.
        /// </summary>
        ErrorMimeContentConversionFailed,

        /// <summary>
        /// Invalid MIME content.
        /// </summary>
        ErrorMimeContentInvalid,

        /// <summary>
        /// Invalid base64 string for MIME content.
        /// </summary>
        ErrorMimeContentInvalidBase64String,

        /// <summary>
        /// The subscription has missed events, but will continue service on this connection.
        /// </summary>
        ErrorMissedNotificationEvents,

        /// <summary>
        /// ErrorMissingArgument
        /// </summary>
        ErrorMissingArgument,

        /// <summary>
        /// When making a request as an account that does not have a mailbox, you must specify the
        /// mailbox primary SMTP address for any distinguished folder Ids.
        /// </summary>
        ErrorMissingEmailAddress,

        /// <summary>
        /// When making a request with an account that does not have a mailbox, you must specify the
        /// primary SMTP address for an existing mailbox.
        /// </summary>
        ErrorMissingEmailAddressForManagedFolder,

        /// <summary>
        /// EmailAddress or ItemId must be included in the request.
        /// </summary>
        ErrorMissingInformationEmailAddress,

        /// <summary>
        /// ReferenceItemId must be included in the request.
        /// </summary>
        ErrorMissingInformationReferenceItemId,

        /// <summary>
        /// SharingFolderId must be included in the request.
        /// </summary>
        ErrorMissingInformationSharingFolderId,

        /// <summary>
        /// An item must be specified when creating an item attachment.
        /// </summary>
        ErrorMissingItemForCreateItemAttachment,

        /// <summary>
        /// The managed folder Id is missing.
        /// </summary>
        ErrorMissingManagedFolderId,

        /// <summary>
        /// A message needs to have at least one recipient.
        /// </summary>
        ErrorMissingRecipients,

        /// <summary>
        /// Missing information for delegate user. You must either specify a valid SMTP address or
        /// SID.
        /// </summary>
        ErrorMissingUserIdInformation,

        /// <summary>
        /// Only one access mode header may be specified.
        /// </summary>
        ErrorMoreThanOneAccessModeSpecified,

        /// <summary>
        /// The move or copy operation failed.
        /// </summary>
        ErrorMoveCopyFailed,

        /// <summary>
        /// Cannot move distinguished folder.
        /// </summary>
        ErrorMoveDistinguishedFolder,

        /// <summary>
        /// ErrorMultiLegacyMailboxAccess
        /// </summary>
        ErrorMultiLegacyMailboxAccess,

        /// <summary>
        /// Multiple results were found.
        /// </summary>
        ErrorNameResolutionMultipleResults,

        /// <summary>
        /// User must have a mailbox for name resolution operations.
        /// </summary>
        ErrorNameResolutionNoMailbox,

        /// <summary>
        /// No results were found.
        /// </summary>
        ErrorNameResolutionNoResults,

        /// <summary>
        /// Another connection was opened against this subscription.
        /// </summary>
        ErrorNewEventStreamConnectionOpened,

        /// <summary>
        /// Exchange Web Services are not currently available for this request because there are no
        /// available Client Access Services Servers in the target AD Site.
        /// </summary>
        ErrorNoApplicableProxyCASServersAvailable,

        /// <summary>
        /// ErrorNoCalendar
        /// </summary>
        ErrorNoCalendar,

        /// <summary>
        /// Exchange Web Services aren't available for this request because there is no Client Access
        /// server with the necessary configuration in the Active Directory site where the mailbox is
        /// stored. If the problem continues, click Help.
        /// </summary>
        ErrorNoDestinationCASDueToKerberosRequirements,

        /// <summary>
        /// Exchange Web Services aren't currently available for this request because an SSL
        /// connection couldn't be established to the Client Access server that should be used for
        /// mailbox access. If the problem continues, click Help.
        /// </summary>
        ErrorNoDestinationCASDueToSSLRequirements,

        /// <summary>
        /// Exchange Web Services aren't currently available for this request because the Client
        /// Access server used for proxying has an older version of Exchange installed than the
        /// Client Access server in the mailbox Active Directory site.
        /// </summary>
        ErrorNoDestinationCASDueToVersionMismatch,

        /// <summary>
        /// You cannot specify the FolderClass when creating a non-generic folder.
        /// </summary>
        ErrorNoFolderClassOverride,

        /// <summary>
        /// ErrorNoFreeBusyAccess
        /// </summary>
        ErrorNoFreeBusyAccess,

        /// <summary>
        /// Mailbox does not exist.
        /// </summary>
        ErrorNonExistentMailbox,

        /// <summary>
        /// The primary SMTP address must be specified when referencing a mailbox.
        /// </summary>
        ErrorNonPrimarySmtpAddress,

        /// <summary>
        /// Custom properties cannot be specified using property tags.  The GUID and Id/Name
        /// combination must be used instead.
        /// </summary>
        ErrorNoPropertyTagForCustomProperties,

        /// <summary>
        /// ErrorNoPublicFolderReplicaAvailable
        /// </summary>
        ErrorNoPublicFolderReplicaAvailable,

        /// <summary>
        /// There are no public folder servers available.
        /// </summary>
        ErrorNoPublicFolderServerAvailable,

        /// <summary>
        /// Exchange Web Services are not currently available for this request because none of the
        /// Client Access Servers in the destination site could process the request.
        /// </summary>
        ErrorNoRespondingCASInDestinationSite,

        /// <summary>
        /// Policy does not allow granting of permissions to external users.
        /// </summary>
        ErrorNotAllowedExternalSharingByPolicy,

        /// <summary>
        /// The user is not a delegate for the mailbox.
        /// </summary>
        ErrorNotDelegate,

        /// <summary>
        /// There was not enough memory to complete the request.
        /// </summary>
        ErrorNotEnoughMemory,

        /// <summary>
        /// The sharing message is not supported.
        /// </summary>
        ErrorNotSupportedSharingMessage,

        /// <summary>
        /// Operation would change object type, which is not permitted.
        /// </summary>
        ErrorObjectTypeChanged,

        /// <summary>
        /// Modified occurrence is crossing or overlapping adjacent occurrence.
        /// </summary>
        ErrorOccurrenceCrossingBoundary,

        /// <summary>
        /// One occurrence of the recurring calendar item overlaps with another occurrence of the
        /// same calendar item.
        /// </summary>
        ErrorOccurrenceTimeSpanTooBig,

        /// <summary>
        /// Operation not allowed with public folder root.
        /// </summary>
        ErrorOperationNotAllowedWithPublicFolderRoot,

        /// <summary>
        /// Organization is not federated.
        /// </summary>
        ErrorOrganizationNotFederated,

        /// <summary>
        /// ErrorOutlookRuleBlobExists
        /// </summary>
        ErrorOutlookRuleBlobExists,

        /// <summary>
        /// You must specify the parent folder Id for this operation.
        /// </summary>
        ErrorParentFolderIdRequired,

        /// <summary>
        /// The specified parent folder could not be found.
        /// </summary>
        ErrorParentFolderNotFound,

        /// <summary>
        /// Password change is required.
        /// </summary>
        ErrorPasswordChangeRequired,

        /// <summary>
        /// Password has expired. Change password.
        /// </summary>
        ErrorPasswordExpired,

        /// <summary>
        /// Policy does not allow granting permission level to user.
        /// </summary>
        ErrorPermissionNotAllowedByPolicy,

        /// <summary>
        /// Dialing restrictions are preventing the phone number that was entered from being dialed.
        /// </summary>
        ErrorPhoneNumberNotDialable,

        /// <summary>
        /// Property update did not succeed.
        /// </summary>
        ErrorPropertyUpdate,

        /// <summary>
        /// At least one property failed validation.
        /// </summary>
        ErrorPropertyValidationFailure,

        /// <summary>
        /// Subscription related request failed because EWS could not contact the appropriate CAS
        /// server for this request.  If this problem persists, recreate the subscription.
        /// </summary>
        ErrorProxiedSubscriptionCallFailure,

        /// <summary>
        /// Request failed because EWS could not contact the appropriate CAS server for this request.
        /// </summary>
        ErrorProxyCallFailed,

        /// <summary>
        /// Exchange Web Services (EWS) is not available for this mailbox because the user account
        /// associated with the mailbox is a member of too many groups. EWS limits the group
        /// membership it can proxy between Client Access Service Servers to 3000.
        /// </summary>
        ErrorProxyGroupSidLimitExceeded,

        /// <summary>
        /// ErrorProxyRequestNotAllowed
        /// </summary>
        ErrorProxyRequestNotAllowed,

        /// <summary>
        /// ErrorProxyRequestProcessingFailed
        /// </summary>
        ErrorProxyRequestProcessingFailed,

        /// <summary>
        /// Exchange Web Services are not currently available for this mailbox because it could not
        /// determine the Client Access Services Server to use for the mailbox.
        /// </summary>
        ErrorProxyServiceDiscoveryFailed,

        /// <summary>
        /// Proxy token has expired.
        /// </summary>
        ErrorProxyTokenExpired,

        /// <summary>
        /// ErrorPublicFolderRequestProcessingFailed
        /// </summary>
        ErrorPublicFolderRequestProcessingFailed,

        /// <summary>
        /// ErrorPublicFolderServerNotFound
        /// </summary>
        ErrorPublicFolderServerNotFound,

        /// <summary>
        /// The search folder has a restriction that is too long to return.
        /// </summary>
        ErrorQueryFilterTooLong,

        /// <summary>
        /// Mailbox has exceeded maximum mailbox size.
        /// </summary>
        ErrorQuotaExceeded,

        /// <summary>
        /// Unable to retrieve events for this subscription.  The subscription must be recreated.
        /// </summary>
        ErrorReadEventsFailed,

        /// <summary>
        /// Unable to suppress read receipt. Read receipts are not pending.
        /// </summary>
        ErrorReadReceiptNotPending,

        /// <summary>
        /// Recurrence end date can not exceed Sep 1, 4500 00:00:00.
        /// </summary>
        ErrorRecurrenceEndDateTooBig,

        /// <summary>
        /// Recurrence has no occurrences in the specified range.
        /// </summary>
        ErrorRecurrenceHasNoOccurrence,

        /// <summary>
        /// Failed to remove one or more delegates.
        /// </summary>
        ErrorRemoveDelegatesFailed,

        /// <summary>
        /// ErrorRequestAborted
        /// </summary>
        ErrorRequestAborted,

        /// <summary>
        /// ErrorRequestStreamTooBig
        /// </summary>
        ErrorRequestStreamTooBig,

        /// <summary>
        /// Required property is missing.
        /// </summary>
        ErrorRequiredPropertyMissing,

        /// <summary>
        /// Cannot perform ResolveNames for non-contact folder.
        /// </summary>
        ErrorResolveNamesInvalidFolderType,

        /// <summary>
        /// Only one contacts folder can be specified in request.
        /// </summary>
        ErrorResolveNamesOnlyOneContactsFolderAllowed,

        /// <summary>
        /// The response failed schema validation.
        /// </summary>
        ErrorResponseSchemaValidation,

        /// <summary>
        /// The restriction or sort order is too complex for this operation.
        /// </summary>
        ErrorRestrictionTooComplex,

        /// <summary>
        /// Restriction contained too many elements.
        /// </summary>
        ErrorRestrictionTooLong,

        /// <summary>
        /// ErrorResultSetTooBig
        /// </summary>
        ErrorResultSetTooBig,

        /// <summary>
        /// ErrorRulesOverQuota
        /// </summary>
        ErrorRulesOverQuota,

        /// <summary>
        /// The folder in which items were to be saved could not be found.
        /// </summary>
        ErrorSavedItemFolderNotFound,

        /// <summary>
        /// The request failed schema validation.
        /// </summary>
        ErrorSchemaValidation,

        /// <summary>
        /// The search folder is not initialized.
        /// </summary>
        ErrorSearchFolderNotInitialized,

        /// <summary>
        /// The user account which was used to submit this request does not have the right to send
        /// mail on behalf of the specified sending account.
        /// </summary>
        ErrorSendAsDenied,

        /// <summary>
        /// SendMeetingCancellations attribute is required for Calendar items.
        /// </summary>
        ErrorSendMeetingCancellationsRequired,

        /// <summary>
        /// The SendMeetingInvitationsOrCancellations attribute is required for calendar items.
        /// </summary>
        ErrorSendMeetingInvitationsOrCancellationsRequired,

        /// <summary>
        /// The SendMeetingInvitations attribute is required for calendar items.
        /// </summary>
        ErrorSendMeetingInvitationsRequired,

        /// <summary>
        /// The meeting request has already been sent and might not be updated.
        /// </summary>
        ErrorSentMeetingRequestUpdate,

        /// <summary>
        /// The task request has already been sent and may not be updated.
        /// </summary>
        ErrorSentTaskRequestUpdate,

        /// <summary>
        /// The server cannot service this request right now. Try again later.
        /// </summary>
        ErrorServerBusy,

        /// <summary>
        /// ErrorServiceDiscoveryFailed
        /// </summary>
        ErrorServiceDiscoveryFailed,

        /// <summary>
        /// No external Exchange Web Service URL available.
        /// </summary>
        ErrorSharingNoExternalEwsAvailable,

        /// <summary>
        /// Failed to synchronize the sharing folder.
        /// </summary>
        ErrorSharingSynchronizationFailed,

        /// <summary>
        /// The current ChangeKey is required for this operation.
        /// </summary>
        ErrorStaleObject,

        /// <summary>
        /// The message couldn't be sent because the sender's submission quota was exceeded. Please
        /// try again later.
        /// </summary>
        ErrorSubmissionQuotaExceeded,

        /// <summary>
        /// Access is denied. Only the subscription owner may access the subscription.
        /// </summary>
        ErrorSubscriptionAccessDenied,

        /// <summary>
        /// Subscriptions are not supported for delegate user access.
        /// </summary>
        ErrorSubscriptionDelegateAccessNotSupported,

        /// <summary>
        /// The specified subscription was not found.
        /// </summary>
        ErrorSubscriptionNotFound,

        /// <summary>
        /// The StreamingSubscription was unsubscribed while the current connection was servicing it.
        /// </summary>
        ErrorSubscriptionUnsubscribed,

        /// <summary>
        /// The folder to be synchronized could not be found.
        /// </summary>
        ErrorSyncFolderNotFound,

        /// <summary>
        /// ErrorTeamMailboxNotFound
        /// </summary>
        ErrorTeamMailboxNotFound,

        /// <summary>
        /// ErrorTeamMailboxNotLinkedToSharePoint
        /// </summary>
        ErrorTeamMailboxNotLinkedToSharePoint,

        /// <summary>
        /// ErrorTeamMailboxUrlValidationFailed
        /// </summary>
        ErrorTeamMailboxUrlValidationFailed,

        /// <summary>
        /// ErrorTeamMailboxNotAuthorizedOwner
        /// </summary>
        ErrorTeamMailboxNotAuthorizedOwner,

        /// <summary>
        /// ErrorTeamMailboxActiveToPendingDelete
        /// </summary>
        ErrorTeamMailboxActiveToPendingDelete,

        /// <summary>
        /// ErrorTeamMailboxFailedSendingNotifications
        /// </summary>
        ErrorTeamMailboxFailedSendingNotifications,

        /// <summary>
        /// ErrorTeamMailboxErrorUnknown
        /// </summary>
        ErrorTeamMailboxErrorUnknown,

        /// <summary>
        /// ErrorTimeIntervalTooBig
        /// </summary>
        ErrorTimeIntervalTooBig,

        /// <summary>
        /// ErrorTimeoutExpired
        /// </summary>
        ErrorTimeoutExpired,

        /// <summary>
        /// The time zone isn't valid.
        /// </summary>
        ErrorTimeZone,

        /// <summary>
        /// The specified target folder could not be found.
        /// </summary>
        ErrorToFolderNotFound,

        /// <summary>
        /// The requesting account does not have permission to serialize tokens.
        /// </summary>
        ErrorTokenSerializationDenied,

        /// <summary>
        /// ErrorUnableToGetUserOofSettings
        /// </summary>
        ErrorUnableToGetUserOofSettings,

        /// <summary>
        /// ErrorUnableToRemoveImContactFromGroup
        /// </summary>
        ErrorUnableToRemoveImContactFromGroup,

        /// <summary>
        /// A dial plan could not be found.
        /// </summary>
        ErrorUnifiedMessagingDialPlanNotFound,

        /// <summary>
        /// The UnifiedMessaging request failed.
        /// </summary>
        ErrorUnifiedMessagingRequestFailed,

        /// <summary>
        /// A connection couldn't be made to the Unified Messaging server.
        /// </summary>
        ErrorUnifiedMessagingServerNotFound,

        /// <summary>
        /// The specified item culture is not supported on this server.
        /// </summary>
        ErrorUnsupportedCulture,

        /// <summary>
        /// The MAPI property type is not supported.
        /// </summary>
        ErrorUnsupportedMapiPropertyType,

        /// <summary>
        /// MIME conversion is not supported for this item type.
        /// </summary>
        ErrorUnsupportedMimeConversion,

        /// <summary>
        /// The property can not be used with this type of restriction.
        /// </summary>
        ErrorUnsupportedPathForQuery,

        /// <summary>
        /// The property can not be used for sorting or grouping results.
        /// </summary>
        ErrorUnsupportedPathForSortGroup,

        /// <summary>
        /// PropertyDefinition is not supported in searches.
        /// </summary>
        ErrorUnsupportedPropertyDefinition,

        /// <summary>
        /// QueryFilter type is not supported.
        /// </summary>
        ErrorUnsupportedQueryFilter,

        /// <summary>
        /// The specified recurrence is not supported.
        /// </summary>
        ErrorUnsupportedRecurrence,

        /// <summary>
        /// Unsupported subfilter type.
        /// </summary>
        ErrorUnsupportedSubFilter,

        /// <summary>
        /// Unsupported type for restriction conversion.
        /// </summary>
        ErrorUnsupportedTypeForConversion,

        /// <summary>
        /// Failed to update one or more delegates.
        /// </summary>
        ErrorUpdateDelegatesFailed,

        /// <summary>
        /// Property for update does not match property in object.
        /// </summary>
        ErrorUpdatePropertyMismatch,

        /// <summary>
        /// Policy does not allow granting permissions to user.
        /// </summary>
        ErrorUserNotAllowedByPolicy,

        /// <summary>
        /// The user isn't enabled for Unified Messaging
        /// </summary>
        ErrorUserNotUnifiedMessagingEnabled,

        /// <summary>
        /// The user doesn't have an SMTP proxy address from a federated domain.
        /// </summary>
        ErrorUserWithoutFederatedProxyAddress,

        /// <summary>
        /// The value is out of range.
        /// </summary>
        ErrorValueOutOfRange,

        /// <summary>
        /// Virus detected in the message.
        /// </summary>
        ErrorVirusDetected,

        /// <summary>
        /// The item has been deleted as a result of a virus scan.
        /// </summary>
        ErrorVirusMessageDeleted,

        /// <summary>
        /// The Voice Mail distinguished folder is not implemented.
        /// </summary>
        ErrorVoiceMailNotImplemented,

        /// <summary>
        /// ErrorWebRequestInInvalidState
        /// </summary>
        ErrorWebRequestInInvalidState,

        /// <summary>
        /// ErrorWin32InteropError
        /// </summary>
        ErrorWin32InteropError,

        /// <summary>
        /// ErrorWorkingHoursSaveFailed
        /// </summary>
        ErrorWorkingHoursSaveFailed,

        /// <summary>
        /// ErrorWorkingHoursXmlMalformed
        /// </summary>
        ErrorWorkingHoursXmlMalformed,

        /// <summary>
        /// The Client Access server version doesn't match the Mailbox server version of the resource
        /// that was being accessed. To determine the correct URL to use to access the resource, use
        /// Autodiscover with the address of the resource.
        /// </summary>
        ErrorWrongServerVersion,

        /// <summary>
        /// The mailbox of the authenticating user and the mailbox of the resource being accessed
        /// must have the same Mailbox server version.
        /// </summary>
        ErrorWrongServerVersionDelegate,

        /// <summary>
        /// The client access token request is invalid.
        /// </summary>
        ErrorInvalidClientAccessTokenRequest,

        /// <summary>
        /// invalid managementrole header value or usage.
        /// </summary>
        ErrorInvalidManagementRoleHeader,

        /// <summary>
        /// SearchMailboxes query has too many keywords.
        /// </summary>
        ErrorSearchQueryHasTooManyKeywords,

        /// <summary>
        /// SearchMailboxes on too many mailboxes.
        /// </summary>
        ErrorSearchTooManyMailboxes,

        /// <summary>There are no retention tags.</summary>
        ErrorInvalidRetentionTagNone,

        /// <summary>Discovery Searches are disabled.</summary>
        ErrorDiscoverySearchesDisabled,

        /// <summary>SeekToConditionPageView not supported for calendar items.</summary>
        ErrorCalendarSeekToConditionNotSupported,

        /// <summary>Archive mailbox search operation failed.</summary>
        ErrorArchiveMailboxSearchFailed,

        /// <summary>Get remote archive mailbox folder failed.</summary>
        ErrorGetRemoteArchiveFolderFailed,

        /// <summary>Find remote archive mailbox folder failed.</summary>
        ErrorFindRemoteArchiveFolderFailed,

        /// <summary>Get remote archive mailbox item failed.</summary>
        ErrorGetRemoteArchiveItemFailed,

        /// <summary>Export remote archive mailbox items failed.</summary>
        ErrorExportRemoteArchiveItemsFailed,

        /// <summary>Invalid state definition.</summary>
        ErrorClientIntentInvalidStateDefinition,

        /// <summary>Client intent not found.</summary>
        ErrorClientIntentNotFound,

        /// <summary>The Content Indexing service is required to perform this search, but it's not enabled.</summary>
        ErrorContentIndexingNotEnabled,

        /// <summary>The custom prompt files you specified couldn't be removed.</summary>
        ErrorDeleteUnifiedMessagingPromptFailed,

        /// <summary>The location service is disabled.</summary>
        ErrorLocationServicesDisabled,

        /// <summary>Invalid location service request.</summary>
        ErrorLocationServicesInvalidRequest,

        /// <summary>The request for location information failed.</summary>
        ErrorLocationServicesRequestFailed,

        /// <summary>The request for location information timed out.</summary>
        ErrorLocationServicesRequestTimedOut,

        /// <summary>Weather service is disabled.</summary>
        ErrorWeatherServiceDisabled,

        /// <summary>Mailbox scope not allowed without a query string.</summary>
        ErrorMailboxScopeNotAllowedWithoutQueryString,

        /// <summary>No speech detected.</summary>
        ErrorNoSpeechDetected,

        /// <summary>An error occurred while accessing the custom prompt publishing point.</summary>
        ErrorPromptPublishingOperationFailed,

        /// <summary>Unable to discover the URL of the public folder mailbox.</summary>
        ErrorPublicFolderMailboxDiscoveryFailed,

        /// <summary>Public folder operation failed.</summary>
        ErrorPublicFolderOperationFailed,

        /// <summary>The operation succeeded on the primary public folder mailbox, but failed to sync to the secondary public folder mailbox.</summary>
        ErrorPublicFolderSyncException,

        /// <summary>Discovery Searches are disabled.</summary>
        ErrorRecipientNotFound,

        /// <summary>Recognizer not installed.</summary>
        ErrorRecognizerNotInstalled,

        /// <summary>Speech grammar error.</summary>
        ErrorSpeechGrammarError,

        /// <summary>Too many concurrent connections opened.</summary>
        ErrorTooManyObjectsOpened,

        /// <summary>Unified Messaging server unavailable.</summary>
        ErrorUMServerUnavailable,

        /// <summary>The Unified Messaging custom prompt file you specified couldn't be found.</summary>
        ErrorUnifiedMessagingPromptNotFound,

        /// <summary>Report data for the UM call summary couldn't be found.</summary>
        ErrorUnifiedMessagingReportDataNotFound,

        /// <summary>The requested size is invalid.</summary>
        ErrorInvalidPhotoSize,

        /// <summary>
        /// AcceptItem action is invalid for a meeting message in group mailbox.
        /// </summary>
        ErrorCalendarIsGroupMailboxForAccept,

        /// <summary>
        /// DeclineItem operation is invalid for a meeting message in group mailbox.
        /// </summary>
        ErrorCalendarIsGroupMailboxForDecline,

        /// <summary>
        /// TentativelyAcceptItem action isn't valid for a meeting message in group mailbox.
        /// </summary>
        ErrorCalendarIsGroupMailboxForTentative,

        /// <summary>
        /// SuppressReadReceipt action isn't valid for a meeting message in group mailbox.
        /// </summary>
        ErrorCalendarIsGroupMailboxForSuppressReadReceipt,

        /// <summary>
        /// The Organization is marked for removal.
        /// </summary>
        ErrorOrganizationAccessBlocked,

        /// <summary>
        /// User doesn't have a valid license.
        /// </summary>
        ErrorInvalidLicense,

        /// <summary>
        /// Receive quota message per folder is exceeded.
        /// </summary>
        ErrorMessagePerFolderCountReceiveQuotaExceeded
    }
}
