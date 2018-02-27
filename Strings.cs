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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Microsoft.Exchange.WebServices
{
    internal static class Strings
    {
        internal static string CannotRemoveSubscriptionFromLiveConnection = "Subscriptions can't be removed from an open connection.";
        internal static string ReadAccessInvalidForNonCalendarFolder = "The Permission read access value {0} can't be used with a non-calendar folder.";
        internal static string PropertyDefinitionPropertyMustBeSet = "The PropertyDefinition property must be set.";
        internal static string ArgumentIsBlankString = "The string argument contains only white space characters.";
        internal static string InvalidAutodiscoverDomainsCount = "At least one domain name must be requested.";
        internal static string MinutesMustBeBetween0And1439 = "minutes must be between 0 and 1439, inclusive.";
        internal static string DeleteInvalidForUnsavedUserConfiguration = "This user configuration object can't be deleted because it's never been saved.";
        internal static string PeriodNotFound = "Invalid transition. A period with the specified Id couldn't be found: {0}";
        internal static string InvalidAutodiscoverSmtpAddress = "A valid SMTP address must be specified.";
        internal static string InvalidOAuthToken = "The given token is invalid.";
        internal static string MaxScpHopsExceeded = "The number of SCP URL hops exceeded the limit.";
        internal static string ContactGroupMemberCannotBeUpdatedWithoutBeingLoadedFirst = "The contact group's Members property must be reloaded before newly-added members can be updated.";
        internal static string CurrentPositionNotElementStart = "The current position is not the start of an element.";
        internal static string CannotConvertBetweenTimeZones = "Unable to convert {0} from {1} to {2}.";
        internal static string FrequencyMustBeBetween1And1440 = "The frequency must be a value between 1 and 1440.";
        internal static string CannotSetDelegateFolderPermissionLevelToCustom = "This operation can't be performed because one or more folder permission levels were set to Custom.";
        internal static string PartnerTokenIncompatibleWithRequestVersion = "TryGetPartnerAccess only supports {0} or a later version in Microsoft-hosted data center.";
        internal static string InvalidAutodiscoverRequest = "Invalid Autodiscover request: '{0}'";
        internal static string InvalidAsyncResult = "The IAsyncResult object was not returned from the corresponding asynchronous method of the original ExchangeService object.";
        internal static string InvalidMailboxType = "The mailbox type isn't valid.";
        internal static string AttachmentCollectionNotLoaded = "The attachment collection must be loaded.";
        internal static string ParameterIncompatibleWithRequestVersion = "The parameter {0} is only valid for Exchange Server version {1} or a later version.";
        internal static string DayOfWeekIndexMustBeSpecifiedForRecurrencePattern = "The recurrence pattern's DayOfWeekIndex property must be specified.";
        internal static string WLIDCredentialsCannotBeUsedWithLegacyAutodiscover = "This type of credentials can't be used with this AutodiscoverService.";
        internal static string PropertyCannotBeUpdated = "This property can't be updated.";
        internal static string IncompatibleTypeForArray = "Type {0} can't be used as an array of type {1}.";
        internal static string PercentCompleteMustBeBetween0And100 = "PercentComplete must be between 0 and 100.";
        internal static string AutodiscoverServiceIncompatibleWithRequestVersion = "The Autodiscover service only supports {0} or a later version.";
        internal static string InvalidAutodiscoverSmtpAddressesCount = "At least one SMTP address must be requested.";
        internal static string ServiceUrlMustBeSet = "The Url property on the ExchangeService object must be set.";
        internal static string ItemTypeNotCompatible = "The item type returned by the service ({0}) isn't compatible with the requested item type ({1}).";
        internal static string AttachmentItemTypeMismatch = "Can not update this attachment item since the item in the response has a different type.";
        internal static string UnsupportedWebProtocol = "Protocol {0} isn't supported for service requests.";
        internal static string EnumValueIncompatibleWithRequestVersion = "Enumeration value {0} in enumeration type {1} is only valid for Exchange version {2} or later.";
        internal static string UnexpectedElement = "An element node '{0}:{1}' of the type {2} was expected, but node '{3}' of type {4} was found.";
        internal static string InvalidOrderBy = "At least one of the property definitions in the OrderBy clause is null.";
        internal static string NoAppropriateConstructorForItemClass = "No appropriate constructor could be found for this item class.";
        internal static string SearchFilterAtIndexIsInvalid = "The search filter at index {0} is invalid.";
        internal static string DeletingThisObjectTypeNotAuthorized = "Deleting this type of object isn't authorized.";
        internal static string PropertyCannotBeDeleted = "This property can't be deleted.";
        internal static string ValuePropertyMustBeSet = "The Value property must be set.";
        internal static string TagValueIsOutOfRange = "The extended property tag value must be in the range of 0 to 65,535.";
        internal static string ItemToUpdateCannotBeNullOrNew = "Items[{0}] is either null or does not have an Id.";
        internal static string SearchParametersRootFolderIdsEmpty = "SearchParameters must contain at least one folder id.";
        internal static string MailboxQueriesParameterIsNotSpecified = "The collection of query and mailboxes parameter is not specified.";
        internal static string FolderPermissionHasInvalidUserId = "The UserId in the folder permission at index {0} is invalid. The StandardUser, PrimarySmtpAddress, or SID property must be set.";
        internal static string InvalidAutodiscoverDomain = "The domain name must be specified.";
        internal static string MailboxesParameterIsNotSpecified = "The array of mailboxes (in legacy DN) is not specified.";
        internal static string ParentFolderDoesNotHaveId = "parentFolder doesn't have an Id.";
        internal static string DayOfMonthMustBeSpecifiedForRecurrencePattern = "The recurrence pattern's DayOfMonth property must be specified.";
        internal static string ClassIncompatibleWithRequestVersion = "Class {0} is only valid for Exchange version {1} or later.";
        internal static string CertificateHasNoPrivateKey = "The given certificate does not have the private key. The private key is necessary to sign part of the request message.";
        internal static string InvalidOrUnsupportedTimeZoneDefinition = "The time zone definition is invalid or unsupported.";
        internal static string HourMustBeBetween0And23 = "Hour must be between 0 and 23.";
        internal static string TimeoutMustBeBetween1And1440 = "Timeout must be a value between 1 and 1440.";
        internal static string CredentialsRequired = "Credentials are required to make a service request.";
        internal static string MustLoadOrAssignPropertyBeforeAccess = "You must load or assign this property before you can read its value.";
        internal static string InvalidAutodiscoverServiceResponse = "The Autodiscover service response was invalid.";
        internal static string CannotCallConnectDuringLiveConnection = "The connection has already opened.";
        internal static string ObjectDoesNotHaveId = "This service object doesn't have an ID.";
        internal static string CannotAddSubscriptionToLiveConnection = "Subscriptions can't be added to an open connection.";
        internal static string MaxChangesMustBeBetween1And512 = "MaxChangesReturned must be between 1 and 512.";
        internal static string AttributeValueCannotBeSerialized = "Values of type '{0}' can't be used for the '{1}' attribute.";
        internal static string NumberOfDaysMustBePositive = "NumberOfDays must be zero or greater. Zero indicates no limit.";
        internal static string SearchFilterMustBeSet = "The SearchFilter property must be set.";
        internal static string EndDateMustBeGreaterThanStartDate = "EndDate must be greater than StartDate.";
        internal static string InvalidDateTime = "Invalid date and time: {0}.";
        internal static string UpdateItemsDoesNotAllowAttachments = "This operation can't be performed because attachments have been added or deleted for one or more items.";
        internal static string TimeoutMustBeGreaterThanZero = "Timeout must be greater than zero.";
        internal static string AutodiscoverInvalidSettingForOutlookProvider = "The requested setting, '{0}', isn't supported by this Autodiscover endpoint.";
        internal static string InvalidRedirectionResponseReturned = "The service returned an invalid redirection response.";
        internal static string ExpectedStartElement = "The start element was expected, but node '{0}' of type {1} was found.";
        internal static string DaysOfTheWeekNotSpecified = "The recurrence pattern's property DaysOfTheWeek must contain at least one day of the week.";
        internal static string FolderToUpdateCannotBeNullOrNew = "Folders[{0}] is either null or does not have an Id.";
        internal static string PartnerTokenRequestRequiresUrl = "TryGetPartnerAccess request requires the Url be set with the partner's autodiscover url first.";
        internal static string NumberOfOccurrencesMustBeGreaterThanZero = "NumberOfOccurrences must be greater than 0.";
        internal static string StartTimeZoneRequired = "StartTimeZone required when setting the Start, End, IsAllDayEvent, or Recurrence properties.  You must load or assign this property before attempting to update the appointment.";
        internal static string PropertyAlreadyExistsInOrderByCollection = "Property {0} already exists in OrderByCollection.";
        internal static string ItemAttachmentMustBeNamed = "The name of the item attachment at index {0} must be set.";
        internal static string InvalidAutodiscoverSettingsCount = "At least one setting must be requested.";
        internal static string LoadingThisObjectTypeNotSupported = "Loading this type of object is not supported.";
        internal static string UserIdForDelegateUserNotSpecified = "The UserId in the DelegateUser hasn't been specified.";
        internal static string PhoneCallAlreadyDisconnected = "The phone call has already been disconnected.";
        internal static string OperationDoesNotSupportAttachments = "This operation isn't supported on attachments.";
        internal static string UnsupportedTimeZonePeriodTransitionTarget = "The time zone transition target isn't supported.";
        internal static string IEnumerableDoesNotContainThatManyObject = "The IEnumerable doesn't contain that many objects.";
        internal static string UpdateItemsDoesNotSupportNewOrUnchangedItems = "This operation can't be performed because one or more items are new or unmodified.";
        internal static string ValidationFailed = "Validation failed.";
        internal static string InvalidRecurrencePattern = "Invalid recurrence pattern: ({0}).";
        internal static string TimeWindowStartTimeMustBeGreaterThanEndTime = "The time window's end time must be greater than its start time.";
        internal static string InvalidAttributeValue = "The invalid value '{0}' was specified for the '{1}' attribute.";
        internal static string FileAttachmentContentIsNotSet = "The content of the file attachment at index {0} must be set.";
        internal static string AutodiscoverDidNotReturnEwsUrl = "The Autodiscover service didn't return an appropriate URL that can be used for the ExchangeService Autodiscover URL.";
        internal static string RecurrencePatternMustHaveStartDate = "The recurrence pattern's StartDate property must be specified.";
        internal static string OccurrenceIndexMustBeGreaterThanZero = "OccurrenceIndex must be greater than 0.";
        internal static string ServiceResponseDoesNotContainXml = "The response received from the service didn't contain valid XML.";
        internal static string ItemIsOutOfDate = "The operation can't be performed because the item is out of date. Reload the item and try again.";
        internal static string MinuteMustBeBetween0And59 = "Minute must be between 0 and 59.";
        internal static string NoSoapOrWsSecurityEndpointAvailable = "No appropriate Autodiscover SOAP or WS-Security endpoint is available.";
        internal static string ElementNotFound = "The element '{0}' in namespace '{1}' wasn't found at the current position.";
        internal static string IndexIsOutOfRange = "index is out of range.";
        internal static string PropertyIsReadOnly = "This property is read-only and can't be set.";
        internal static string AttachmentCreationFailed = "At least one attachment couldn't be created.";
        internal static string DayOfMonthMustBeBetween1And31 = "DayOfMonth must be between 1 and 31.";
        internal static string ServiceRequestFailed = "The request failed. {0}";
        internal static string DelegateUserHasInvalidUserId = "The UserId in the DelegateUser is invalid. The StandardUser, PrimarySmtpAddress or SID property must be set.";
        internal static string SearchFilterComparisonValueTypeIsNotSupported = "Values of type '{0}' can't be used as comparison values in search filters.";
        internal static string ElementValueCannotBeSerialized = "Values of type '{0}' can't be used for the '{1}' element.";
        internal static string PropertyValueMustBeSpecifiedForRecurrencePattern = "The recurrence pattern's {0} property must be specified.";
        internal static string NonSummaryPropertyCannotBeUsed = "The property {0} can't be used in {1} requests.";
        internal static string HoldIdParameterIsNotSpecified = "The hold id parameter is not specified.";
        internal static string TransitionGroupNotFound = "Invalid transition. A transition group with the specified ID couldn't be found: {0}";
        internal static string ObjectTypeNotSupported = "Objects of type {0} can't be added to the dictionary. The following types are supported: string array, byte array, boolean, byte, DateTime, integer, long, string, unsigned integer, and unsigned long.";
        internal static string InvalidTimeoutValue = "{0} is not a valid timeout value. Valid values range from 1 to 1440.";
        internal static string AutodiscoverRedirectBlocked = "Autodiscover blocked a potentially insecure redirection to {0}. To allow Autodiscover to follow the redirection, use the AutodiscoverUrl(string, AutodiscoverRedirectionUrlValidationCallback) overload.";
        internal static string PropertySetCannotBeModified = "This PropertySet is read-only and can't be modified.";
        internal static string DayOfTheWeekMustBeSpecifiedForRecurrencePattern = "The recurrence pattern's property DayOfTheWeek must be specified.";
        internal static string ServiceObjectAlreadyHasId = "This operation can't be performed because this service object already has an ID. To update this service object, use the Update() method instead.";
        internal static string MethodIncompatibleWithRequestVersion = "Method {0} is only valid for Exchange Server version {1} or later.";
        internal static string OperationNotSupportedForPropertyDefinitionType = "This operation isn't supported for property definition type {0}.";
        internal static string InvalidElementStringValue = "The invalid value '{0}' was specified for the '{1}' element.";
        internal static string CollectionIsEmpty = "The collection is empty.";
        internal static string InvalidFrequencyValue = "{0} is not a valid frequency value. Valid values range from 1 to 1440.";
        internal static string UnexpectedEndOfXmlDocument = "The XML document ended unexpectedly.";
        internal static string FolderTypeNotCompatible = "The folder type returned by the service ({0}) isn't compatible with the requested folder type ({1}).";
        internal static string RequestIncompatibleWithRequestVersion = "The service request {0} is only valid for Exchange version {1} or later.";
        internal static string PropertyTypeIncompatibleWhenUpdatingCollection = "Can not update the existing collection item since the item in the response has a different type.";
        internal static string ServerVersionNotSupported = "Exchange Server doesn't support the requested version.";
        internal static string DurationMustBeSpecifiedWhenScheduled = "Duration must be specified when State is equal to Scheduled.";
        internal static string NoError = "No error.";
        internal static string CannotUpdateNewUserConfiguration = "This user configuration can't be updated because it's never been saved.";
        internal static string ObjectTypeIncompatibleWithRequestVersion = "The object type {0} is only valid for Exchange Server version {1} or later versions.";
        internal static string NullStringArrayElementInvalid = "The array contains at least one null element.";
        internal static string HttpsIsRequired = "Https is required when partner token is expected.";
        internal static string MergedFreeBusyIntervalMustBeSmallerThanTimeWindow = "MergedFreeBusyInterval must be smaller than the specified time window.";
        internal static string SecondMustBeBetween0And59 = "Second must be between 0 and 59.";
        internal static string AtLeastOneAttachmentCouldNotBeDeleted = "At least one attachment couldn't be deleted.";
        internal static string IdAlreadyInList = "The ID is already in the list.";
        internal static string BothSearchFilterAndQueryStringCannotBeSpecified = "Both search filter and query string can't be specified. One of them must be null.";
        internal static string AdditionalPropertyIsNull = "The additional property at index {0} is null.";
        internal static string InvalidEmailAddress = "The e-mail address is formed incorrectly.";
        internal static string MaximumRedirectionHopsExceeded = "The maximum redirection hop count has been reached.";
        internal static string AutodiscoverCouldNotBeLocated = "The Autodiscover service couldn't be located.";
        internal static string NoSubscriptionsOnConnection = "You must add at least one subscription to this connection before it can be opened.";
        internal static string PermissionLevelInvalidForNonCalendarFolder = "The Permission level value {0} can't be used with a non-calendar folder.";
        internal static string InvalidAuthScheme = "The token auth scheme should be bearer.";
        internal static string ValuePropertyNotLoaded = "This property was requested, but it wasn't returned by the server.";
        internal static string PropertyIncompatibleWithRequestVersion = "The property {0} is valid only for Exchange {1} or later versions.";
        internal static string OffsetMustBeGreaterThanZero = "The offset must be greater than 0.";
        internal static string CreateItemsDoesNotAllowAttachments = "This operation doesn't support items that have attachments.";
        internal static string PropertyDefinitionTypeMismatch = "Property definition type '{0}' and type parameter '{1}' aren't compatible.";
        internal static string IntervalMustBeGreaterOrEqualToOne = "The interval must be greater than or equal to 1.";
        internal static string CannotSetPermissionLevelToCustom = "The PermissionLevel property can't be set to FolderPermissionLevel.Custom. To define a custom permission, set its individual properties to the values you want.";
        internal static string CannotAddRequestHeader = "HTTP header '{0}' isn't permitted. Only HTTP headers with the 'X-' prefix are permitted.";
        internal static string ArrayMustHaveAtLeastOneElement = "The Array value must have at least one element.";
        internal static string MonthMustBeSpecifiedForRecurrencePattern = "The recurrence pattern's Month property must be specified.";
        internal static string ValueOfTypeCannotBeConverted = "The value '{0}' of type {1} can't be converted to a value of type {2}.";
        internal static string ValueCannotBeConverted = "The value '{0}' couldn't be converted to type {1}.";
        internal static string ServerErrorAndStackTraceDetails = "{0} -- Server Error: {1}: {2} {3}";
        internal static string FolderPermissionLevelMustBeSet = "The permission level of the folder permission at index {0} must be set.";
        internal static string AutodiscoverError = "The Autodiscover service returned an error.";
        internal static string ArrayMustHaveSingleDimension = "The array value must have a single dimension.";
        internal static string InvalidPropertyValueNotInRange = "{0} must be between {1} and {2}.";
        internal static string RegenerationPatternsOnlyValidForTasks = "Regeneration patterns can only be used with Task items.";
        internal static string ItemAttachmentCannotBeUpdated = "Item attachments can't be updated.";
        internal static string EqualityComparisonFilterIsInvalid = "Either the OtherPropertyDefinition or the Value properties must be set.";
        internal static string AutodiscoverServiceRequestRequiresDomainOrUrl = "This Autodiscover request requires that either the Domain or Url be specified.";
        internal static string InvalidUser = "Invalid user: '{0}'";
        internal static string AccountIsLocked = "This account is locked. Visit {0} to unlock it.";
        internal static string InvalidDomainName = "'{0}' is not a valid domain name.";
        internal static string TooFewServiceReponsesReturned = "The service was expected to return {1} responses of type '{0}', but {2} responses were received.";
        internal static string CannotSubscribeToStatusEvents = "Status events can't be subscribed to.";
        internal static string InvalidSortByPropertyForMailboxSearch = "Specified SortBy property '{0}' is invalid.";
        internal static string UnexpectedElementType = "The expected XML node type was {0}, but the actual type is {1}.";
        internal static string ValueMustBeGreaterThanZero = "The value must be greater than 0.";
        internal static string AttachmentCannotBeUpdated = "Attachments can't be updated.";
        internal static string CreateItemsDoesNotHandleExistingItems = "This operation can't be performed because at least one item already has an ID.";
        internal static string MultipleContactPhotosInAttachment = "This operation only allows at most 1 file attachment with IsContactPhoto set.";
        internal static string InvalidRecurrenceRange = "Invalid recurrence range: ({0}).";
        internal static string CannotSetBothImpersonatedAndPrivilegedUser = "Can't set both impersonated user and privileged user in the ExchangeService object.";
        internal static string NewMessagesWithAttachmentsCannotBeSentDirectly = "New messages with attachments can't be sent directly. You must first save the message and then send it.";
        internal static string CannotCallDisconnectWithNoLiveConnection = "The connection is already closed.";
        internal static string IdPropertyMustBeSet = "The Id property must be set.";
        internal static string ValuePropertyNotAssigned = "You must assign this property before you can read its value.";
        internal static string ZeroLengthArrayInvalid = "The array must contain at least one element.";
        internal static string HoldMailboxesParameterIsNotSpecified = "The hold mailboxes parameter is not specified.";
        internal static string CannotSaveNotNewUserConfiguration = "Calling Save isn't allowed because this user configuration isn't new. To apply local changes to this user configuration, call Update instead.";
        internal static string ServiceObjectDoesNotHaveId = "This operation can't be performed because this service object doesn't have an Id.";
        internal static string PropertyCollectionSizeMismatch = "The collection returned by the service has a different size from the current one.";
        internal static string XsDurationCouldNotBeParsed = "The specified xsDuration argument couldn't be parsed.";
        internal static string UnknownTimeZonePeriodTransitionType = "Unknown time zone transition type: {0}";
        internal static string UserPhotoSizeNotSpecified = "The UserPhotoSize must be not be null or empty.";
        internal static string UserPhotoNotSpecified = "The photo must be not be null or empty.";

    }
}