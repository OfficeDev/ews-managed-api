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
    /// Defines the available properties of a rule. 
    /// </summary>
    public enum RuleProperty
    {
        /// <summary>
        /// The RuleId property of a rule.
        /// </summary>
        [EwsEnum("RuleId")]
        RuleId,

        /// <summary>
        /// The DisplayName property of a rule.
        /// </summary>
        [EwsEnum("DisplayName")]
        DisplayName,

        /// <summary>
        /// The Priority property of a rule.
        /// </summary>
        [EwsEnum("Priority")]
        Priority,

        /// <summary>
        /// The IsNotSupported property of a rule.
        /// </summary>
        [EwsEnum("IsNotSupported")]
        IsNotSupported,

        /// <summary>
        /// The Actions property of a rule.
        /// </summary>
        [EwsEnum("Actions")]
        Actions,

        /// <summary>
        /// The Categories property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:Categories")]
        ConditionCategories,

        /// <summary>
        /// The ContainsBodyStrings property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:ContainsBodyStrings")]
        ConditionContainsBodyStrings,

        /// <summary>
        /// The ContainsHeaderStrings property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:ContainsHeaderStrings")]
        ConditionContainsHeaderStrings,

        /// <summary>
        /// The ContainsRecipientStrings property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:ContainsRecipientStrings")]
        ConditionContainsRecipientStrings,

        /// <summary>
        /// The ContainsSenderStrings property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:ContainsSenderStrings")]
        ConditionContainsSenderStrings,

        /// <summary>
        /// The ContainsSubjectOrBodyStrings property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:ContainsSubjectOrBodyStrings")]
        ConditionContainsSubjectOrBodyStrings,

        /// <summary>
        /// The ContainsSubjectStrings property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:ContainsSubjectStrings")]
        ConditionContainsSubjectStrings,

        /// <summary>
        /// The FlaggedForAction property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:FlaggedForAction")]
        ConditionFlaggedForAction,

        /// <summary>
        /// The FromAddresses property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:FromAddresses")]
        ConditionFromAddresses,

        /// <summary>
        /// The FromConnectedAccounts property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:FromConnectedAccounts")]
        ConditionFromConnectedAccounts,

        /// <summary>
        /// The HasAttachments property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:HasAttachments")]
        ConditionHasAttachments,

        /// <summary>
        /// The Importance property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:Importance")]
        ConditionImportance,

        /// <summary>
        /// The IsApprovalRequest property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:IsApprovalRequest")]
        ConditionIsApprovalRequest,

        /// <summary>
        /// The IsAutomaticForward property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:IsAutomaticForward")]
        ConditionIsAutomaticForward,

        /// <summary>
        /// The IsAutomaticReply property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:IsAutomaticReply")]
        ConditionIsAutomaticReply,

        /// <summary>
        /// The IsEncrypted property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:IsEncrypted")]
        ConditionIsEncrypted,

        /// <summary>
        /// The IsMeetingRequest property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:IsMeetingRequest")]
        ConditionIsMeetingRequest,

        /// <summary>
        /// The IsMeetingResponse property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:IsMeetingResponse")]
        ConditionIsMeetingResponse,

        /// <summary>
        /// The IsNonDeliveryReport property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:IsNDR")]
        ConditionIsNonDeliveryReport,

        /// <summary>
        /// The IsPermissionControlled property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:IsPermissionControlled")]
        ConditionIsPermissionControlled,

        /// <summary>
        /// The IsRead property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:IsRead")]
        ConditionIsRead,

        /// <summary>
        /// The IsSigned property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:IsSigned")]
        ConditionIsSigned,

        /// <summary>
        /// The IsVoicemail property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:IsVoicemail")]
        ConditionIsVoicemail,

        /// <summary>
        /// The IsReadReceipt property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:IsReadReceipt")]
        ConditionIsReadReceipt,

        /// <summary>
        /// The ItemClasses property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:ItemClasses")]
        ConditionItemClasses,

        /// <summary>
        /// The MessageClassifications property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:MessageClassifications")]
        ConditionMessageClassifications,

        /// <summary>
        /// The NotSentToMe property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:NotSentToMe")]
        ConditionNotSentToMe,

        /// <summary>
        /// The SentCcMe property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:SentCcMe")]
        ConditionSentCcMe,

        /// <summary>
        /// The SentOnlyToMe property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:SentOnlyToMe")]
        ConditionSentOnlyToMe,

        /// <summary>
        /// The SentToAddresses property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:SentToAddresses")]
        ConditionSentToAddresses,

        /// <summary>
        /// The SentToMe property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:SentToMe")]
        ConditionSentToMe,

        /// <summary>
        /// The SentToOrCcMe property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:SentToOrCcMe")]
        ConditionSentToOrCcMe,

        /// <summary>
        /// The Sensitivity property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:Sensitivity")]
        ConditionSensitivity,

        /// <summary>
        /// The WithinDateRange property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:WithinDateRange")]
        ConditionWithinDateRange,

        /// <summary>
        /// The WithinSizeRange property of a rule's set of conditions.
        /// </summary>
        [EwsEnum("Condition:WithinSizeRange")]
        ConditionWithinSizeRange,

        /// <summary>
        /// The Categories property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:Categories")]
        ExceptionCategories,

        /// <summary>
        /// The ContainsBodyStrings property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:ContainsBodyStrings")]
        ExceptionContainsBodyStrings,

        /// <summary>
        /// The ContainsHeaderStrings property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:ContainsHeaderStrings")]
        ExceptionContainsHeaderStrings,

        /// <summary>
        /// The ContainsRecipientStrings property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:ContainsRecipientStrings")]
        ExceptionContainsRecipientStrings,

        /// <summary>
        /// The ContainsSenderStrings property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:ContainsSenderStrings")]
        ExceptionContainsSenderStrings,

        /// <summary>
        /// The ContainsSubjectOrBodyStrings property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:ContainsSubjectOrBodyStrings")]
        ExceptionContainsSubjectOrBodyStrings,

        /// <summary>
        /// The ContainsSubjectStrings property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:ContainsSubjectStrings")]
        ExceptionContainsSubjectStrings,

        /// <summary>
        /// The FlaggedForAction property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:FlaggedForAction")]
        ExceptionFlaggedForAction,

        /// <summary>
        /// The FromAddresses property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:FromAddresses")]
        ExceptionFromAddresses,

        /// <summary>
        /// The FromConnectedAccounts property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:FromConnectedAccounts")]
        ExceptionFromConnectedAccounts,

        /// <summary>
        /// The HasAttachments property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:HasAttachments")]
        ExceptionHasAttachments,

        /// <summary>
        /// The Importance property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:Importance")]
        ExceptionImportance,

        /// <summary>
        /// The IsApprovalRequest property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:IsApprovalRequest")]
        ExceptionIsApprovalRequest,

        /// <summary>
        /// The IsAutomaticForward property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:IsAutomaticForward")]
        ExceptionIsAutomaticForward,

        /// <summary>
        /// The IsAutomaticReply property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:IsAutomaticReply")]
        ExceptionIsAutomaticReply,

        /// <summary>
        /// The IsEncrypted property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:IsEncrypted")]
        ExceptionIsEncrypted,

        /// <summary>
        /// The IsMeetingRequest property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:IsMeetingRequest")]
        ExceptionIsMeetingRequest,

        /// <summary>
        /// The IsMeetingResponse property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:IsMeetingResponse")]
        ExceptionIsMeetingResponse,

        /// <summary>
        /// The IsNonDeliveryReport property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:IsNDR")]
        ExceptionIsNonDeliveryReport,

        /// <summary>
        /// The IsPermissionControlled property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:IsPermissionControlled")]
        ExceptionIsPermissionControlled,

        /// <summary>
        /// The IsRead property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:IsRead")]
        ExceptionIsRead,

        /// <summary>
        /// The IsSigned property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:IsSigned")]
        ExceptionIsSigned,

        /// <summary>
        /// The IsVoicemail property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:IsVoicemail")]
        ExceptionIsVoicemail,

        /// <summary>
        /// The ItemClasses property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:ItemClasses")]
        ExceptionItemClasses,

        /// <summary>
        /// The MessageClassifications property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:MessageClassifications")]
        ExceptionMessageClassifications,

        /// <summary>
        /// The NotSentToMe property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:NotSentToMe")]
        ExceptionNotSentToMe,

        /// <summary>
        /// The SentCcMe property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:SentCcMe")]
        ExceptionSentCcMe,

        /// <summary>
        /// The SentOnlyToMe property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:SentOnlyToMe")]
        ExceptionSentOnlyToMe,

        /// <summary>
        /// The SentToAddresses property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:SentToAddresses")]
        ExceptionSentToAddresses,

        /// <summary>
        /// The SentToMe property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:SentToMe")]
        ExceptionSentToMe,

        /// <summary>
        /// The SentToOrCcMe property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:SentToOrCcMe")]
        ExceptionSentToOrCcMe,

        /// <summary>
        /// The Sensitivity property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:Sensitivity")]
        ExceptionSensitivity,

        /// <summary>
        /// The WithinDateRange property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:WithinDateRange")]
        ExceptionWithinDateRange,

        /// <summary>
        /// The WithinSizeRange property of a rule's set of exceptions.
        /// </summary>
        [EwsEnum("Exception:WithinSizeRange")]
        ExceptionWithinSizeRange,

        /// <summary>
        /// The Categories property in a rule's set of actions.
        /// </summary>
        [EwsEnum("Action:Categories")]
        ActionCategories,

        /// <summary>
        /// The CopyToFolder property in a rule's set of actions.
        /// </summary>
        [EwsEnum("Action:CopyToFolder")]
        ActionCopyToFolder,

        /// <summary>
        /// The Delete property in a rule's set of actions.
        /// </summary>
        [EwsEnum("Action:Delete")]
        ActionDelete,

        /// <summary>
        /// The ForwardAsAttachmentToRecipients property in a rule's set of actions.
        /// </summary>
        [EwsEnum("Action:ForwardAsAttachmentToRecipients")]
        ActionForwardAsAttachmentToRecipients,

        /// <summary>
        /// The ForwardToRecipients property in a rule's set of actions.
        /// </summary>
        [EwsEnum("Action:ForwardToRecipients")]
        ActionForwardToRecipients,

        /// <summary>
        /// The Importance property in a rule's set of actions.
        /// </summary>
        [EwsEnum("Action:Importance")]
        ActionImportance,

        /// <summary>
        /// The MarkAsRead property in a rule's set of actions.
        /// </summary>
        [EwsEnum("Action:MarkAsRead")]
        ActionMarkAsRead,

        /// <summary>
        /// The MoveToFolder property in a rule's set of actions.
        /// </summary>
        [EwsEnum("Action:MoveToFolder")]
        ActionMoveToFolder,

        /// <summary>
        /// The PermanentDelete property in a rule's set of actions.
        /// </summary>
        [EwsEnum("Action:PermanentDelete")]
        ActionPermanentDelete,

        /// <summary>
        /// The RedirectToRecipients property in a rule's set of actions.
        /// </summary>
        [EwsEnum("Action:RedirectToRecipients")]
        ActionRedirectToRecipients,

        /// <summary>
        /// The SendSMSAlertToRecipients property in a rule's set of actions.
        /// </summary>
        [EwsEnum("Action:SendSMSAlertToRecipients")]
        ActionSendSMSAlertToRecipients,

        /// <summary>
        /// The ServerReplyWithMessage property in a rule's set of actions.
        /// </summary>
        [EwsEnum("Action:ServerReplyWithMessage")]
        ActionServerReplyWithMessage,

        /// <summary>
        /// The StopProcessingRules property in a rule's set of actions.
        /// </summary>
        [EwsEnum("Action:StopProcessingRules")]
        ActionStopProcessingRules,

        /// <summary>
        /// The IsEnabled property of a rule, indicating if the rule is enabled.
        /// </summary>
        [EwsEnum("IsEnabled")]
        IsEnabled,

        /// <summary>
        /// The IsInError property of a rule, indicating if the rule is in error.
        /// </summary>
        [EwsEnum("IsInError")]
        IsInError,

        /// <summary>
        /// The Conditions property of a rule, contains all conditions of the rule.
        /// </summary>
        [EwsEnum("Conditions")]
        Conditions,

        /// <summary>
        /// The Exceptions property of a rule, contains all exceptions of the rule.
        /// </summary>
        [EwsEnum("Exceptions")]
        Exceptions
    }
}