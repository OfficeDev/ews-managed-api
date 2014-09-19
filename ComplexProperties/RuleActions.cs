// ---------------------------------------------------------------------------
// <copyright file="RuleActions.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the RuleActions class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.ObjectModel;

    /// <summary>
    /// Represents the set of actions available for a rule.
    /// </summary>
    public sealed class RuleActions : ComplexProperty
    {
        /// <summary>
        /// SMS recipient address type.
        /// </summary>
        private const string MobileType = "MOBILE";

        /// <summary>
        /// The AssignCategories action.
        /// </summary>
        private StringList assignCategories;

        /// <summary>
        /// The CopyToFolder action.
        /// </summary>
        private FolderId copyToFolder;

        /// <summary>
        /// The Delete action.
        /// </summary>
        private bool delete;

        /// <summary>
        /// The ForwardAsAttachmentToRecipients action.
        /// </summary>
        private EmailAddressCollection forwardAsAttachmentToRecipients;

        /// <summary>
        /// The ForwardToRecipients action.
        /// </summary>
        private EmailAddressCollection forwardToRecipients;

        /// <summary>
        /// The MarkImportance action.
        /// </summary>
        private Importance? markImportance;

        /// <summary>
        /// The MarkAsRead action.
        /// </summary>
        private bool markAsRead;

        /// <summary>
        /// The MoveToFolder action.
        /// </summary>
        private FolderId moveToFolder;

        /// <summary>
        /// The PermanentDelete action.
        /// </summary>
        private bool permanentDelete;

        /// <summary>
        /// The RedirectToRecipients action.
        /// </summary>
        private EmailAddressCollection redirectToRecipients;

        /// <summary>
        /// The SendSMSAlertToRecipients action.
        /// </summary>
        private Collection<MobilePhone> sendSMSAlertToRecipients;

        /// <summary>
        /// The ServerReplyWithMessage action.
        /// </summary>
        private ItemId serverReplyWithMessage;

        /// <summary>
        /// The StopProcessingRules action.
        /// </summary>
        private bool stopProcessingRules;

        /// <summary>
        /// Initializes a new instance of the <see cref="RulePredicates"/> class.
        /// </summary>
        internal RuleActions()
            : base()
        {
            this.assignCategories = new StringList();
            this.forwardAsAttachmentToRecipients = new EmailAddressCollection(XmlElementNames.Address);
            this.forwardToRecipients = new EmailAddressCollection(XmlElementNames.Address);
            this.redirectToRecipients = new EmailAddressCollection(XmlElementNames.Address);
            this.sendSMSAlertToRecipients = new Collection<MobilePhone>();
        }

        /// <summary>
        /// Gets the categories that should be stamped on incoming messages. 
        /// To disable stamping incoming messages with categories, set 
        /// AssignCategories to null.
        /// </summary>
        public StringList AssignCategories
        {
            get
            {
                return this.assignCategories;
            }
        }

        /// <summary>
        /// Gets or sets the Id of the folder incoming messages should be copied to.
        /// To disable copying incoming messages to a folder, set CopyToFolder to null.
        /// </summary>
        public FolderId CopyToFolder
        {
            get
            {
                return this.copyToFolder;
            }

            set
            {
                this.SetFieldValue<FolderId>(ref this.copyToFolder, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages should be
        /// automatically moved to the Deleted Items folder.
        /// </summary>
        public bool Delete
        {
            get
            {
                return this.delete;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.delete, value);
            }
        }

        /// <summary>
        /// Gets the e-mail addresses to which incoming messages should be 
        /// forwarded as attachments. To disable forwarding incoming messages
        /// as attachments, empty the ForwardAsAttachmentToRecipients list.
        /// </summary>
        public EmailAddressCollection ForwardAsAttachmentToRecipients
        {
            get
            {
                return this.forwardAsAttachmentToRecipients;
            }
        }

        /// <summary>
        /// Gets the e-mail addresses to which incoming messages should be forwarded. 
        /// To disable forwarding incoming messages, empty the ForwardToRecipients list.
        /// </summary>
        public EmailAddressCollection ForwardToRecipients
        {
            get
            {
                return this.forwardToRecipients;
            }
        }

        /// <summary>
        /// Gets or sets the importance that should be stamped on incoming 
        /// messages. To disable the stamping of incoming messages with an 
        /// importance, set MarkImportance to null.
        /// </summary>
        public Importance? MarkImportance
        {
            get
            {
                return this.markImportance;
            }

            set
            {
                this.SetFieldValue<Importance?>(ref this.markImportance, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages should be 
        /// marked as read.
        /// </summary>
        public bool MarkAsRead
        {
            get
            {
                return this.markAsRead;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.markAsRead, value);
            }
        }

        /// <summary>
        /// Gets or sets the Id of the folder to which incoming messages should be
        /// moved. To disable the moving of incoming messages to a folder, set
        /// CopyToFolder to null.
        /// </summary>
        public FolderId MoveToFolder
        {
            get
            {
                return this.moveToFolder;
            }

            set
            {
                this.SetFieldValue<FolderId>(ref this.moveToFolder, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages should be 
        /// permanently deleted. When a message is permanently deleted, it is never 
        /// saved into the recipient's mailbox. To delete a message after it has 
        /// been saved into the recipient's mailbox, use the Delete action.
        /// </summary>
        public bool PermanentDelete
        {
            get
            {
                return this.permanentDelete;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.permanentDelete, value);
            }
        }

        /// <summary>
        /// Gets the e-mail addresses to which incoming messages should be 
        /// redirecteded. To disable redirection of incoming messages, empty
        /// the RedirectToRecipients list. Unlike forwarded mail, redirected mail
        /// maintains the original sender and recipients. 
        /// </summary>
        public EmailAddressCollection RedirectToRecipients
        {
            get
            {
                return this.redirectToRecipients;
            }
        }

        /// <summary>
        /// Gets the phone numbers to which an SMS alert should be sent. To disable
        /// sending SMS alerts for incoming messages, empty the 
        /// SendSMSAlertToRecipients list.
        /// </summary>
        public Collection<MobilePhone> SendSMSAlertToRecipients
        {
            get
            {
                return this.sendSMSAlertToRecipients;
            }
        }

        /// <summary>
        /// Gets or sets the Id of the template message that should be sent
        /// as a reply to incoming messages. To disable automatic replies, set 
        /// ServerReplyWithMessage to null. 
        /// </summary>
        public ItemId ServerReplyWithMessage
        {
            get
            {
                return this.serverReplyWithMessage;
            }

            set
            {
                this.SetFieldValue<ItemId>(ref this.serverReplyWithMessage, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether subsequent rules should be
        /// evaluated. 
        /// </summary>
        public bool StopProcessingRules
        {
            get
            {
                return this.stopProcessingRules;
            }

            set
            {
                this.SetFieldValue<bool>(ref this.stopProcessingRules, value);
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
                case XmlElementNames.AssignCategories:
                    this.assignCategories.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.CopyToFolder:
                    reader.ReadStartElement(XmlNamespace.NotSpecified, XmlElementNames.FolderId);
                    this.copyToFolder = new FolderId();
                    this.copyToFolder.LoadFromXml(reader, XmlElementNames.FolderId);
                    reader.ReadEndElement(XmlNamespace.NotSpecified, XmlElementNames.CopyToFolder);
                    return true;
                case XmlElementNames.Delete:
                    this.delete = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.ForwardAsAttachmentToRecipients:
                    this.forwardAsAttachmentToRecipients.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.ForwardToRecipients:
                    this.forwardToRecipients.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.MarkImportance:
                    this.markImportance = reader.ReadElementValue<Importance>();
                    return true;
                case XmlElementNames.MarkAsRead:
                    this.markAsRead = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.MoveToFolder:
                    reader.ReadStartElement(XmlNamespace.NotSpecified, XmlElementNames.FolderId);
                    this.moveToFolder = new FolderId();
                    this.moveToFolder.LoadFromXml(reader, XmlElementNames.FolderId);
                    reader.ReadEndElement(XmlNamespace.NotSpecified, XmlElementNames.MoveToFolder);
                    return true;
                case XmlElementNames.PermanentDelete:
                    this.permanentDelete = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.RedirectToRecipients:
                    this.redirectToRecipients.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.SendSMSAlertToRecipients:
                    EmailAddressCollection smsRecipientCollection = new EmailAddressCollection(XmlElementNames.Address);
                    smsRecipientCollection.LoadFromXml(reader, reader.LocalName);
                    this.sendSMSAlertToRecipients = ConvertSMSRecipientsFromEmailAddressCollectionToMobilePhoneCollection(smsRecipientCollection);
                    return true;
                case XmlElementNames.ServerReplyWithMessage:
                    this.serverReplyWithMessage = new ItemId();
                    this.serverReplyWithMessage.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.StopProcessingRules:
                    this.stopProcessingRules = reader.ReadElementValue<bool>();
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlElementNames.AssignCategories:
                        this.assignCategories.LoadFromJson(jsonProperty.ReadAsJsonObject(key), service);
                        break;
                    case XmlElementNames.CopyToFolder:
                        this.copyToFolder = new FolderId();
                        this.copyToFolder.LoadFromJson(
                            jsonProperty.ReadAsJsonObject(key), 
                            service);
                        break;
                    case XmlElementNames.Delete:
                        this.delete = jsonProperty.ReadAsBool(key);
                        break;
                    case XmlElementNames.ForwardAsAttachmentToRecipients:
                        this.forwardAsAttachmentToRecipients.LoadFromJson(jsonProperty.ReadAsJsonObject(key), service);
                        break;
                    case XmlElementNames.ForwardToRecipients:
                        this.forwardToRecipients.LoadFromJson(jsonProperty.ReadAsJsonObject(key), service);
                        break;
                    case XmlElementNames.MarkImportance:
                        this.markImportance = jsonProperty.ReadEnumValue<Importance>(key);
                        break;
                    case XmlElementNames.MarkAsRead:
                        this.markAsRead = jsonProperty.ReadAsBool(key);
                        break;
                    case XmlElementNames.MoveToFolder:
                        this.moveToFolder = new FolderId();
                        this.moveToFolder.LoadFromJson(
                            jsonProperty.ReadAsJsonObject(key), 
                            service);
                        break;
                    case XmlElementNames.PermanentDelete:
                        this.permanentDelete = jsonProperty.ReadAsBool(key);
                        break;
                    case XmlElementNames.RedirectToRecipients:
                        this.redirectToRecipients.LoadFromJson(jsonProperty.ReadAsJsonObject(key), service);
                        break;
                    case XmlElementNames.SendSMSAlertToRecipients:
                        EmailAddressCollection smsRecipientCollection = new EmailAddressCollection(XmlElementNames.Address);
                        smsRecipientCollection.LoadFromJson(jsonProperty.ReadAsJsonObject(key), service);
                        this.sendSMSAlertToRecipients = ConvertSMSRecipientsFromEmailAddressCollectionToMobilePhoneCollection(smsRecipientCollection);
                        break;
                    case XmlElementNames.ServerReplyWithMessage:
                        this.serverReplyWithMessage = new ItemId();
                        this.serverReplyWithMessage.LoadFromJson(jsonProperty.ReadAsJsonObject(key), service);
                        break;
                    case XmlElementNames.StopProcessingRules:
                        this.stopProcessingRules = jsonProperty.ReadAsBool(key);
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (this.AssignCategories.Count > 0)
            {
                this.AssignCategories.WriteToXml(writer, XmlElementNames.AssignCategories);
            }

            if (this.CopyToFolder != null)
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.CopyToFolder);
                this.CopyToFolder.WriteToXml(writer);
                writer.WriteEndElement();
            }

            if (this.Delete != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.Delete, 
                    this.Delete);
            }

            if (this.ForwardAsAttachmentToRecipients.Count > 0)
            {
                this.ForwardAsAttachmentToRecipients.WriteToXml(writer, XmlElementNames.ForwardAsAttachmentToRecipients);
            }

            if (this.ForwardToRecipients.Count > 0)
            {
                this.ForwardToRecipients.WriteToXml(writer, XmlElementNames.ForwardToRecipients);
            }

            if (this.MarkImportance.HasValue)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.MarkImportance, 
                    this.MarkImportance.Value);
            }

            if (this.MarkAsRead != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.MarkAsRead,
                    this.MarkAsRead);
            }

            if (this.MoveToFolder != null)
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.MoveToFolder);
                this.MoveToFolder.WriteToXml(writer);
                writer.WriteEndElement();
            }

            if (this.PermanentDelete != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.PermanentDelete, 
                    this.PermanentDelete);
            }

            if (this.RedirectToRecipients.Count > 0)
            {
                this.RedirectToRecipients.WriteToXml(writer, XmlElementNames.RedirectToRecipients);
            }

            if (this.SendSMSAlertToRecipients.Count > 0)
            {
                EmailAddressCollection emailCollection = ConvertSMSRecipientsFromMobilePhoneCollectionToEmailAddressCollection(this.SendSMSAlertToRecipients);
                emailCollection.WriteToXml(writer, XmlElementNames.SendSMSAlertToRecipients);
            }

            if (this.ServerReplyWithMessage != null)
            {
                this.ServerReplyWithMessage.WriteToXml(writer, XmlElementNames.ServerReplyWithMessage);
            }

            if (this.StopProcessingRules != false)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.StopProcessingRules, 
                    this.StopProcessingRules);
            }
        }

        /// <summary>
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object InternalToJson(ExchangeService service)
        {
            JsonObject jsonProperty = new JsonObject();

            if (this.AssignCategories.Count > 0)
            {
                jsonProperty.Add(XmlElementNames.AssignCategories, this.AssignCategories.InternalToJson(service));
            }

            if (this.CopyToFolder != null)
            {
                jsonProperty.Add(XmlElementNames.CopyToFolder, this.CopyToFolder.InternalToJson(service));
            }

            if (this.Delete != false)
            {
                jsonProperty.Add(XmlElementNames.Delete, this.Delete);
            }

            if (this.ForwardAsAttachmentToRecipients.Count > 0)
            {
                jsonProperty.Add(XmlElementNames.ForwardAsAttachmentToRecipients, this.ForwardAsAttachmentToRecipients.InternalToJson(service));
            }

            if (this.ForwardToRecipients.Count > 0)
            {
                jsonProperty.Add(XmlElementNames.ForwardToRecipients, this.ForwardToRecipients.InternalToJson(service));
            }

            if (this.MarkImportance.HasValue)
            {
                jsonProperty.Add(XmlElementNames.MarkImportance, this.MarkImportance.Value);
            }

            if (this.MarkAsRead != false)
            {
                jsonProperty.Add(XmlElementNames.MarkAsRead, this.MarkAsRead);
            }

            if (this.MoveToFolder != null)
            {
                jsonProperty.Add(XmlElementNames.MoveToFolder, this.MoveToFolder.InternalToJson(service));
            }

            if (this.PermanentDelete != false)
            {
                jsonProperty.Add(XmlElementNames.PermanentDelete, this.PermanentDelete);
            }

            if (this.RedirectToRecipients.Count > 0)
            {
                jsonProperty.Add(XmlElementNames.RedirectToRecipients, this.RedirectToRecipients.InternalToJson(service));
            }

            if (this.SendSMSAlertToRecipients.Count > 0)
            {
                EmailAddressCollection emailCollection = ConvertSMSRecipientsFromMobilePhoneCollectionToEmailAddressCollection(this.SendSMSAlertToRecipients);
                jsonProperty.Add(XmlElementNames.SendSMSAlertToRecipients, emailCollection.InternalToJson(service));
            }

            if (this.ServerReplyWithMessage != null)
            {
                jsonProperty.Add(XmlElementNames.ServerReplyWithMessage, this.ServerReplyWithMessage.InternalToJson(service));
            }

            if (this.StopProcessingRules != false)
            {
                jsonProperty.Add(XmlElementNames.StopProcessingRules, this.StopProcessingRules);
            }

            return jsonProperty;
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void InternalValidate()
        {
            base.InternalValidate();
            EwsUtilities.ValidateParam(this.forwardAsAttachmentToRecipients, "ForwardAsAttachmentToRecipients");
            EwsUtilities.ValidateParam(this.forwardToRecipients, "ForwardToRecipients");
            EwsUtilities.ValidateParam(this.redirectToRecipients, "RedirectToRecipients");
            foreach (MobilePhone sendSMSAlertToRecipient in this.sendSMSAlertToRecipients)
            {
                EwsUtilities.ValidateParam(sendSMSAlertToRecipient, "SendSMSAlertToRecipient");
            }
        }

        /// <summary>
        /// Convert the SMS recipient list from EmailAddressCollection type to MobilePhone collection type.
        /// </summary>
        /// <param name="emailCollection">Recipient list in EmailAddressCollection type.</param>
        /// <returns>A MobilePhone collection object containing all SMS recipient in MobilePhone type. </returns>
        private static Collection<MobilePhone> ConvertSMSRecipientsFromEmailAddressCollectionToMobilePhoneCollection(EmailAddressCollection emailCollection)
        {
            Collection<MobilePhone> mobilePhoneCollection = new Collection<MobilePhone>();
            foreach (EmailAddress emailAddress in emailCollection)
            {
                mobilePhoneCollection.Add(new MobilePhone(emailAddress.Name, emailAddress.Address));
            }

            return mobilePhoneCollection;
        }

        /// <summary>
        /// Convert the SMS recipient list from MobilePhone collection type to EmailAddressCollection type.
        /// </summary>
        /// <param name="recipientCollection">Recipient list in a MobilePhone collection type.</param>
        /// <returns>An EmailAddressCollection object containing recipients with "MOBILE" address type. </returns>
        private static EmailAddressCollection ConvertSMSRecipientsFromMobilePhoneCollectionToEmailAddressCollection(Collection<MobilePhone> recipientCollection)
        {
            EmailAddressCollection emailCollection = new EmailAddressCollection(XmlElementNames.Address);
            foreach (MobilePhone recipient in recipientCollection)
            {
                EmailAddress emailAddress = new EmailAddress(
                    recipient.Name, 
                    recipient.PhoneNumber, 
                    RuleActions.MobileType);
                emailCollection.Add(emailAddress);
            }

            return emailCollection;
        }
    }
}
