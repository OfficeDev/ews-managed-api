// ---------------------------------------------------------------------------
// <copyright file="IconIndex.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the icon index enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// The values indicate the types of item icons to display.
    /// </summary>
    public enum IconIndex
    {
        /// <summary>
        /// A default icon.
        /// </summary>
        Default,

        /// <summary>
        /// Post Item
        /// </summary>
        PostItem,

        /// <summary>
        /// Icon read
        /// </summary>
        MailRead,

        /// <summary>
        /// Icon unread
        /// </summary>
        MailUnread,

        /// <summary>
        /// Icon replied
        /// </summary>
        MailReplied,

        /// <summary>
        /// Icon forwarded
        /// </summary>
        MailForwarded,

        /// <summary>
        /// Icon encrypted
        /// </summary>
        MailEncrypted,

        /// <summary>
        /// Icon S/MIME signed
        /// </summary>
        MailSmimeSigned,

        /// <summary>
        /// Icon encrypted replied
        /// </summary>
        MailEncryptedReplied,

        /// <summary>
        /// Icon S/MIME signed replied
        /// </summary>
        MailSmimeSignedReplied,

        /// <summary>
        /// Icon encrypted forwarded
        /// </summary>
        MailEncryptedForwarded,

        /// <summary>
        /// Icon S/MIME signed forwarded
        /// </summary>
        MailSmimeSignedForwarded,

        /// <summary>
        /// Icon encrypted read
        /// </summary>
        MailEncryptedRead,

        /// <summary>
        /// Icon S/MIME signed read
        /// </summary>
        MailSmimeSignedRead,

        /// <summary>
        /// IRM-protected mail
        /// </summary>
        MailIrm,

        /// <summary>
        /// IRM-protected mail forwarded
        /// </summary>
        MailIrmForwarded,

        /// <summary>
        /// IRM-protected mail replied
        /// </summary>
        MailIrmReplied,

        /// <summary>
        /// Icon sms routed to external messaging system
        /// </summary>
        SmsSubmitted,

        /// <summary>
        /// Icon sms routed to external messaging system
        /// </summary>
        SmsRoutedToDeliveryPoint,

        /// <summary>
        /// Icon sms routed to external messaging system
        /// </summary>
        SmsRoutedToExternalMessagingSystem,

        /// <summary>
        /// Icon sms routed to external messaging system
        /// </summary>
        SmsDelivered,

        /// <summary>
        /// Outlook Default for Contacts
        /// </summary>
        OutlookDefaultForContacts,

        /// <summary>
        /// Icon appointment item
        /// </summary>
        AppointmentItem,

        /// <summary>
        /// Icon appointment recur
        /// </summary>
        AppointmentRecur,

        /// <summary>
        /// Icon appointment meet
        /// </summary>
        AppointmentMeet,

        /// <summary>
        /// Icon appointment meet recur
        /// </summary>
        AppointmentMeetRecur,

        /// <summary>
        /// Icon appointment meet NY
        /// </summary>
        AppointmentMeetNY,

        /// <summary>
        /// Icon appointment meet yes
        /// </summary>
        AppointmentMeetYes,

        /// <summary>
        /// Icon appointment meet no
        /// </summary>
        AppointmentMeetNo,

        /// <summary>
        /// Icon appointment meet maybe
        /// </summary>
        AppointmentMeetMaybe,

        /// <summary>
        /// Icon appointment meet cancel
        /// </summary>
        AppointmentMeetCancel,

        /// <summary>
        /// Icon appointment meet info
        /// </summary>
        AppointmentMeetInfo,

        /// <summary>
        /// Icon task item
        /// </summary>
        TaskItem,

        /// <summary>
        /// Icon task recur
        /// </summary>
        TaskRecur,

        /// <summary>
        /// Icon task owned
        /// </summary>
        TaskOwned,

        /// <summary>
        /// Icon task delegated
        /// </summary>
        TaskDelegated,
    }
}
