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
    /// Represents the set of conditions and exceptions available for a rule.
    /// </summary>
    public sealed class RulePredicates : ComplexProperty
    {
        /// <summary>
        /// The HasCategories predicate.
        /// </summary>
        private StringList categories;

        /// <summary>
        /// The ContainsBodyStrings predicate.
        /// </summary>
        private StringList containsBodyStrings;

        /// <summary>
        /// The ContainsHeaderStrings predicate.
        /// </summary>
        private StringList containsHeaderStrings;

        /// <summary>
        /// The ContainsRecipientStrings predicate.
        /// </summary>
        private StringList containsRecipientStrings;

        /// <summary>
        /// The ContainsSenderStrings predicate.
        /// </summary>
        private StringList containsSenderStrings;

        /// <summary>
        /// The ContainsSubjectOrBodyStrings predicate.
        /// </summary>
        private StringList containsSubjectOrBodyStrings;

        /// <summary>
        /// The ContainsSubjectStrings predicate.
        /// </summary>
        private StringList containsSubjectStrings;

        /// <summary>
        /// The FlaggedForAction predicate.
        /// </summary>
        private FlaggedForAction? flaggedForAction;

        /// <summary>
        /// The FromAddresses predicate.
        /// </summary>
        private EmailAddressCollection fromAddresses;

        /// <summary>
        /// The FromConnectedAccounts predicate.
        /// </summary>
        private StringList fromConnectedAccounts;

        /// <summary>
        /// The HasAttachments predicate.
        /// </summary>
        private bool hasAttachments;

        /// <summary>
        /// The Importance predicate.
        /// </summary>
        private Importance? importance;

        /// <summary>
        /// The IsApprovalRequest predicate.
        /// </summary>
        private bool isApprovalRequest;
        
        /// <summary>
        /// The IsAutomaticForward predicate.
        /// </summary>
        private bool isAutomaticForward;

        /// <summary>
        /// The IsAutomaticReply predicate.
        /// </summary>
        private bool isAutomaticReply;
        
        /// <summary>
        /// The IsEncrypted predicate.
        /// </summary>
        private bool isEncrypted;
        
        /// <summary>
        /// The IsMeetingRequest predicate.
        /// </summary>
        private bool isMeetingRequest;
        
        /// <summary>
        /// The IsMeetingResponse predicate.
        /// </summary>
        private bool isMeetingResponse;
        
        /// <summary>
        /// The IsNDR predicate.
        /// </summary>
        private bool isNonDeliveryReport;
        
        /// <summary>
        /// The IsPermissionControlled predicate.
        /// </summary>
        private bool isPermissionControlled;
        
        /// <summary>
        /// The IsSigned predicate.
        /// </summary>
        private bool isSigned;
        
        /// <summary>
        /// The IsVoicemail predicate.
        /// </summary>
        private bool isVoicemail;
        
        /// <summary>
        /// The IsReadReceipt  predicate.
        /// </summary>
        private bool isReadReceipt;
        
        /// <summary>
        /// ItemClasses predicate.
        /// </summary>
        private StringList itemClasses;

        /// <summary>
        /// The MessageClassifications predicate.
        /// </summary>
        private StringList messageClassifications;
        
        /// <summary>
        /// The NotSentToMe predicate.
        /// </summary>
        private bool notSentToMe;
        
        /// <summary>
        /// SentCcMe predicate.
        /// </summary>
        private bool sentCcMe;
        
        /// <summary>
        /// The SentOnlyToMe predicate.
        /// </summary>
        private bool sentOnlyToMe;
        
        /// <summary>
        /// The SentToAddresses predicate.
        /// </summary>
        private EmailAddressCollection sentToAddresses;

        /// <summary>
        /// The SentToMe predicate.
        /// </summary>
        private bool sentToMe;

        /// <summary>
        /// The SentToOrCcMe predicate.
        /// </summary>
        private bool sentToOrCcMe;

        /// <summary>
        /// The Sensitivity predicate.
        /// </summary>
        private Sensitivity? sensitivity;

        /// <summary>
        /// The Sensitivity predicate.
        /// </summary>
        private RulePredicateDateRange withinDateRange;

        /// <summary>
        /// The Sensitivity predicate.
        /// </summary>
        private RulePredicateSizeRange withinSizeRange;

        /// <summary>
        /// Initializes a new instance of the <see cref="RulePredicates"/> class.
        /// </summary>
        internal RulePredicates()
            : base()
        {
            this.categories = new StringList();
            this.containsBodyStrings = new StringList();
            this.containsHeaderStrings = new StringList();
            this.containsRecipientStrings = new StringList();
            this.containsSenderStrings = new StringList();
            this.containsSubjectOrBodyStrings = new StringList();
            this.containsSubjectStrings = new StringList();
            this.fromAddresses = new EmailAddressCollection(XmlElementNames.Address);
            this.fromConnectedAccounts = new StringList();
            this.itemClasses = new StringList();
            this.messageClassifications = new StringList();
            this.sentToAddresses = new EmailAddressCollection(XmlElementNames.Address);
            this.withinDateRange = new RulePredicateDateRange();
            this.withinSizeRange = new RulePredicateSizeRange();
        }

        /// <summary>
        /// Gets the categories that an incoming message should be stamped with 
        /// for the condition or exception to apply. To disable this predicate,
        /// empty the list.
        /// </summary>
        public StringList Categories
        {
            get
            {
                return this.categories;
            }
        }

        /// <summary>
        /// Gets the strings that should appear in the body of incoming messages 
        /// for the condition or exception to apply.
        /// To disable this predicate, empty the list.
        /// </summary>
        public StringList ContainsBodyStrings
        {
            get
            {
                return this.containsBodyStrings;
            }
        }

        /// <summary>
        /// Gets the strings that should appear in the headers of incoming messages 
        /// for the condition or exception to apply. To disable this predicate, empty 
        /// the list.
        /// </summary>
        public StringList ContainsHeaderStrings
        {
            get
            {
                return this.containsHeaderStrings;
            }
        }

        /// <summary>
        /// Gets the strings that should appear in either the To or Cc fields of 
        /// incoming messages for the condition or exception to apply. To disable this
        /// predicate, empty the list.
        /// </summary>
        public StringList ContainsRecipientStrings
        {
            get
            {
                return this.containsRecipientStrings;
            }
        }

        /// <summary>
        /// Gets the strings that should appear in the From field of incoming messages 
        /// for the condition or exception to apply. To disable this predicate, empty 
        /// the list.
        /// </summary>
        public StringList ContainsSenderStrings
        {
            get
            {
                return this.containsSenderStrings;
            }
        }

        /// <summary>
        /// Gets the strings that should appear in either the body or the subject 
        /// of incoming messages for the condition or exception to apply.
        /// To disable this predicate, empty the list.
        /// </summary>
        public StringList ContainsSubjectOrBodyStrings
        {
            get
            {
                return this.containsSubjectOrBodyStrings;
            }
        }

        /// <summary>
        /// Gets the strings that should appear in the subject of incoming messages 
        /// for the condition or exception to apply. To disable this predicate, 
        /// empty the list.
        /// </summary>
        public StringList ContainsSubjectStrings
        {
            get
            {
                return this.containsSubjectStrings;
            }
        }

        /// <summary>
        /// Gets or sets the flag for action value that should appear on incoming 
        /// messages for the condition or execption to apply. To disable this 
        /// predicate, set it to null. 
        /// </summary>
        public FlaggedForAction? FlaggedForAction
        {
            get
            {
                return this.flaggedForAction;
            }

            set
            {
                this.SetFieldValue<FlaggedForAction?>(ref this.flaggedForAction, value);
            }
        }

        /// <summary>
        /// Gets the e-mail addresses of the senders of incoming messages for the 
        /// condition or exception to apply. To disable this predicate, empty the 
        /// list.
        /// </summary>
        public EmailAddressCollection FromAddresses
        {
            get
            {
                return this.fromAddresses;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must have
        /// attachments for the condition or exception to apply.  
        /// </summary>
        public bool HasAttachments
        {
            get
            {
                return this.hasAttachments;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.hasAttachments, value);
            }
        }

        /// <summary>
        /// Gets or sets the importance that should be stamped on incoming messages 
        /// for the condition or exception to apply. To disable this predicate, set 
        /// it to null.
        /// </summary>
        public Importance? Importance
        {
            get
            {
                return this.importance;
            }

            set
            {
                this.SetFieldValue<Importance?>(ref this.importance, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// approval requests for the condition or exception to apply. 
        /// </summary>
        public bool IsApprovalRequest
        {
            get
            {
                return this.isApprovalRequest;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.isApprovalRequest, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// automatic forwards for the condition or exception to apply.
        /// </summary>
        public bool IsAutomaticForward
        {
            get
            {
                return this.isAutomaticForward;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.isAutomaticForward, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// automatic replies for the condition or exception to apply. 
        /// </summary>
        public bool IsAutomaticReply
        {
            get
            {
                return this.isAutomaticReply;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.isAutomaticReply, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// S/MIME encrypted for the condition or exception to apply.
        /// </summary>
        public bool IsEncrypted
        {
            get
            {
                return this.isEncrypted;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.isEncrypted, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// meeting requests for the condition or exception to apply. 
        /// </summary>
        public bool IsMeetingRequest
        {
            get
            {
                return this.isMeetingRequest;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.isMeetingRequest, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// meeting responses for the condition or exception to apply. 
        /// </summary>
        public bool IsMeetingResponse
        {
            get
            {
                return this.isMeetingResponse;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.isMeetingResponse, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// non-delivery reports (NDR) for the condition or exception to apply. 
        /// </summary>
        public bool IsNonDeliveryReport
        {
            get
            {
                return this.isNonDeliveryReport;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.isNonDeliveryReport, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// permission controlled (RMS protected) for the condition or exception 
        /// to apply. 
        /// </summary>
        public bool IsPermissionControlled
        {
            get
            {
                return this.isPermissionControlled;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.isPermissionControlled, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// S/MIME signed for the condition or exception to apply. 
        /// </summary>
        public bool IsSigned
        {
            get
            {
                return this.isSigned;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.isSigned, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// voice mails for the condition or exception to apply. 
        /// </summary>
        public bool IsVoicemail
        {
            get
            {
                return this.isVoicemail;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.isVoicemail, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// read receipts for the condition or exception to apply. 
        /// </summary>
        public bool IsReadReceipt
        {
            get
            {
                return this.isReadReceipt;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.isReadReceipt, value);
            }
        }

        /// <summary>
        /// Gets the e-mail account names from which incoming messages must have 
        /// been aggregated for the condition or exception to apply. To disable 
        /// this predicate, empty the list.
        /// </summary>
        public StringList FromConnectedAccounts
        {
            get
            {
                return this.fromConnectedAccounts;
            }
        }

        /// <summary>
        /// Gets the item classes that must be stamped on incoming messages for
        /// the condition or exception to apply. To disable this predicate, 
        /// empty the list.
        /// </summary>
        public StringList ItemClasses
        {
            get
            {
                return this.itemClasses;
            }
        }
        
        /// <summary>
        /// Gets the message classifications that must be stamped on incoming messages
        /// for the condition or exception to apply. To disable this predicate, 
        /// empty the list.
        /// </summary>
        public StringList MessageClassifications
        {
            get
            {
                return this.messageClassifications;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the owner of the mailbox must 
        /// NOT be a To recipient of the incoming messages for the condition or 
        /// exception to apply.
        /// </summary>
        public bool NotSentToMe
        {
            get
            {
                return this.notSentToMe;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.notSentToMe, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the owner of the mailbox must be 
        /// a Cc recipient of incoming messages for the condition or exception to apply. 
        /// </summary>
        public bool SentCcMe
        {
            get
            {
                return this.sentCcMe;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.sentCcMe, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the owner of the mailbox must be 
        /// the only To recipient of incoming messages for the condition or exception 
        /// to apply.
        /// </summary>
        public bool SentOnlyToMe
        {
            get
            {
                return this.sentOnlyToMe;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.sentOnlyToMe, value);
            }
        }

        /// <summary>
        /// Gets the e-mail addresses incoming messages must have been sent to for 
        /// the condition or exception to apply. To disable this predicate, empty 
        /// the list.
        /// </summary>
        public EmailAddressCollection SentToAddresses
        {
            get
            {
                return this.sentToAddresses;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the owner of the mailbox must be 
        /// a To recipient of incoming messages for the condition or exception to apply. 
        /// </summary>
        public bool SentToMe
        {
            get
            {
                return this.sentToMe;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.sentToMe, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the owner of the mailbox must be
        /// either a To or Cc recipient of incoming messages for the condition or
        /// exception to apply. 
        /// </summary>
        public bool SentToOrCcMe
        {
            get
            {
                return this.sentToOrCcMe;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.sentToOrCcMe, value);
            }
        }

        /// <summary>
        /// Gets or sets the sensitivity that must be stamped on incoming messages 
        /// for the condition or exception to apply. To disable this predicate, set it
        /// to null.
        /// </summary>
        public Sensitivity? Sensitivity
        {
            get
            {
                return this.sensitivity;
            }

            set
            {
                this.SetFieldValue<Sensitivity?>(ref this.sensitivity, value);
            }
        }

        /// <summary>
        /// Gets the date range within which incoming messages must have been received 
        /// for the condition or exception to apply. To disable this predicate, set both 
        /// its Start and End properties to null.
        /// </summary>
        public RulePredicateDateRange WithinDateRange
        {
            get
            {
                return this.withinDateRange;
            }
        }

        /// <summary>
        /// Gets the minimum and maximum sizes incoming messages must have for the 
        /// condition or exception to apply. To disable this predicate, set both its 
        /// MinimumSize and MaximumSize properties to null.
        /// </summary>
        public RulePredicateSizeRange WithinSizeRange
        {
            get
            {
                return this.withinSizeRange;
            }
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
                case XmlElementNames.Categories:
                    this.categories.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.ContainsBodyStrings:
                    this.containsBodyStrings.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.ContainsHeaderStrings:
                    this.containsHeaderStrings.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.ContainsRecipientStrings:
                    this.containsRecipientStrings.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.ContainsSenderStrings:
                    this.containsSenderStrings.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.ContainsSubjectOrBodyStrings:
                    this.containsSubjectOrBodyStrings.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.ContainsSubjectStrings:
                    this.containsSubjectStrings.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.FlaggedForAction:
                    this.flaggedForAction = reader.ReadElementValue<FlaggedForAction>();
                    return true;
                case XmlElementNames.FromAddresses:
                    this.fromAddresses.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.FromConnectedAccounts:
                    this.fromConnectedAccounts.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.HasAttachments:
                    this.hasAttachments = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.Importance:
                    this.importance = reader.ReadElementValue<Importance>();
                    return true;
                case XmlElementNames.IsApprovalRequest:
                    this.isApprovalRequest = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsAutomaticForward:
                    this.isAutomaticForward = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsAutomaticReply:
                    this.isAutomaticReply = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsEncrypted:
                    this.isEncrypted = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsMeetingRequest:
                    this.isMeetingRequest = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsMeetingResponse:
                    this.isMeetingResponse = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsNDR:
                    this.isNonDeliveryReport = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsPermissionControlled:
                    this.isPermissionControlled = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsSigned:
                    this.isSigned = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsVoicemail:
                    this.isVoicemail = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsReadReceipt:
                    this.isReadReceipt = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.ItemClasses:
                    this.itemClasses.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.MessageClassifications:
                    this.messageClassifications.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.NotSentToMe:
                    this.notSentToMe = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.SentCcMe:
                    this.sentCcMe = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.SentOnlyToMe:
                    this.sentOnlyToMe = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.SentToAddresses:
                    this.sentToAddresses.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.SentToMe:
                    this.sentToMe = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.SentToOrCcMe:
                    this.sentToOrCcMe = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.Sensitivity:
                    this.sensitivity = reader.ReadElementValue<Sensitivity>();
                    return true;
                case XmlElementNames.WithinDateRange:
                    this.withinDateRange.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.WithinSizeRange:
                    this.withinSizeRange.LoadFromXml(reader, reader.LocalName);
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (this.Categories.Count > 0)
            {
                this.Categories.WriteToXml(writer, XmlElementNames.Categories);
            }

            if (this.ContainsBodyStrings.Count > 0)
            {
                this.ContainsBodyStrings.WriteToXml(writer, XmlElementNames.ContainsBodyStrings);
            }

            if (this.ContainsHeaderStrings.Count > 0)
            {
                this.ContainsHeaderStrings.WriteToXml(writer, XmlElementNames.ContainsHeaderStrings);
            }

            if (this.ContainsRecipientStrings.Count > 0)
            {
                this.ContainsRecipientStrings.WriteToXml(writer, XmlElementNames.ContainsRecipientStrings);
            }

            if (this.ContainsSenderStrings.Count > 0)
            {
                this.ContainsSenderStrings.WriteToXml(writer, XmlElementNames.ContainsSenderStrings);
            }

            if (this.ContainsSubjectOrBodyStrings.Count > 0)
            {
                this.ContainsSubjectOrBodyStrings.WriteToXml(writer, XmlElementNames.ContainsSubjectOrBodyStrings);
            }

            if (this.ContainsSubjectStrings.Count > 0)
            {
                this.ContainsSubjectStrings.WriteToXml(writer, XmlElementNames.ContainsSubjectStrings);
            }

            if (this.FlaggedForAction.HasValue)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.FlaggedForAction,
                    this.FlaggedForAction.Value);
            }

            if (this.FromAddresses.Count > 0)
            {
                this.FromAddresses.WriteToXml(writer, XmlElementNames.FromAddresses);
            }

            if (this.FromConnectedAccounts.Count > 0)
            {
                this.FromConnectedAccounts.WriteToXml(writer, XmlElementNames.FromConnectedAccounts);
            }

            if (this.HasAttachments != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.HasAttachments,
                    this.HasAttachments);
            }

            if (this.Importance.HasValue)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.Importance,
                    this.Importance.Value);
            }

            if (this.IsApprovalRequest != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.IsApprovalRequest,
                    this.IsApprovalRequest);
            }

            if (this.IsAutomaticForward != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.IsAutomaticForward,
                    this.IsAutomaticForward);
            }

            if (this.IsAutomaticReply != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.IsAutomaticReply, 
                    this.IsAutomaticReply);
            }

            if (this.IsEncrypted != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.IsEncrypted, 
                    this.IsEncrypted);
            }

            if (this.IsMeetingRequest != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.IsMeetingRequest, 
                    this.IsMeetingRequest);
            }

            if (this.IsMeetingResponse != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.IsMeetingResponse, 
                    this.IsMeetingResponse);
            }

            if (this.IsNonDeliveryReport != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.IsNDR,
                    this.IsNonDeliveryReport);
            }

            if (this.IsPermissionControlled != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.IsPermissionControlled, 
                    this.IsPermissionControlled);
            }

            if (this.isReadReceipt != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.IsReadReceipt,
                    this.IsReadReceipt);
            }

            if (this.IsSigned != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.IsSigned, 
                    this.IsSigned);
            }

            if (this.IsVoicemail != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.IsVoicemail, 
                    this.IsVoicemail);
            }

            if (this.ItemClasses.Count > 0)
            {
                this.ItemClasses.WriteToXml(writer, XmlElementNames.ItemClasses);
            }

            if (this.MessageClassifications.Count > 0)
            {
                this.MessageClassifications.WriteToXml(writer, XmlElementNames.MessageClassifications);
            }

            if (this.NotSentToMe != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.NotSentToMe, 
                    this.NotSentToMe);
            }

            if (this.SentCcMe != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.SentCcMe, 
                    this.SentCcMe);
            }

            if (this.SentOnlyToMe != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.SentOnlyToMe, 
                    this.SentOnlyToMe);
            }

            if (this.SentToAddresses.Count > 0)
            {
                this.SentToAddresses.WriteToXml(writer, XmlElementNames.SentToAddresses);
            }

            if (this.SentToMe != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.SentToMe, 
                    this.SentToMe);
            }

            if (this.SentToOrCcMe != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.SentToOrCcMe, 
                    this.SentToOrCcMe);
            }

            if (this.Sensitivity.HasValue)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.Sensitivity, 
                    this.Sensitivity.Value);
            }

            if (this.WithinDateRange.Start.HasValue || this.WithinDateRange.End.HasValue)
            {
                this.WithinDateRange.WriteToXml(writer, XmlElementNames.WithinDateRange);
            }

            if (this.WithinSizeRange.MaximumSize.HasValue || this.WithinSizeRange.MinimumSize.HasValue)
            {
                this.WithinSizeRange.WriteToXml(writer, XmlElementNames.WithinSizeRange);
            }
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void InternalValidate()
        {
            base.InternalValidate();
            EwsUtilities.ValidateParam(this.fromAddresses, "FromAddresses");
            EwsUtilities.ValidateParam(this.sentToAddresses, "SentToAddresses");
            EwsUtilities.ValidateParam(this.withinDateRange, "WithinDateRange");
            EwsUtilities.ValidateParam(this.withinSizeRange, "WithinSizeRange");
        }
    }
}