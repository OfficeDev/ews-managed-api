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
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Reflection;

    internal delegate object CreateServiceObjectWithServiceParam(ExchangeService srv);

    internal delegate object CreateServiceObjectWithAttachmentParam(ItemAttachment itemAttachment, bool isNew);

    /// <summary>
    /// ServiceObjectInfo contains metadata on how to map from an element name to a ServiceObject type
    /// as well as how to map from a ServiceObject type to appropriate constructors.
    /// </summary>
    internal class ServiceObjectInfo
    {
        private Dictionary<string, Type> xmlElementNameToServiceObjectClassMap;
        private Dictionary<Type, CreateServiceObjectWithServiceParam> serviceObjectConstructorsWithServiceParam;
        private Dictionary<Type, CreateServiceObjectWithAttachmentParam> serviceObjectConstructorsWithAttachmentParam;

        /// <summary>
        /// Default constructor
        /// </summary>
        internal ServiceObjectInfo()
        {
            this.xmlElementNameToServiceObjectClassMap = new Dictionary<string, Type>();
            this.serviceObjectConstructorsWithServiceParam = new Dictionary<Type, CreateServiceObjectWithServiceParam>();
            this.serviceObjectConstructorsWithAttachmentParam = new Dictionary<Type, CreateServiceObjectWithAttachmentParam>();

            this.InitializeServiceObjectClassMap();
        }

        /// <summary>
        /// Initializes the service object class map.
        /// </summary>
        /// <remarks>
        /// If you add a new ServiceObject subclass that can be returned by the Server, add the type
        /// to the class map as well as associated delegate(s) to call the constructor(s).
        /// </remarks>
        private void InitializeServiceObjectClassMap()
        {
            // Appointment
            this.AddServiceObjectType(
                XmlElementNames.CalendarItem,
                typeof(Appointment),
                delegate(ExchangeService srv) { return new Appointment(srv); },
                delegate(ItemAttachment itemAttachment, bool isNew) { return new Appointment(itemAttachment, isNew); });

            // Booking
            this.AddServiceObjectType(
                XmlElementNames.Booking,
                typeof(Booking),
                delegate (ExchangeService srv) { return new Booking(srv); },
                delegate (ItemAttachment itemAttachment, bool isNew) { return new Appointment(itemAttachment, isNew); });

            // CalendarFolder
            this.AddServiceObjectType(
                XmlElementNames.CalendarFolder,
                typeof(CalendarFolder),
                delegate(ExchangeService srv) { return new CalendarFolder(srv); },
                null);

            // Contact
            this.AddServiceObjectType(
                XmlElementNames.Contact,
                typeof(Contact),
                delegate(ExchangeService srv) { return new Contact(srv); },
                delegate(ItemAttachment itemAttachment, bool isNew) { return new Contact(itemAttachment); });

            // ContactsFolder
            this.AddServiceObjectType(
                XmlElementNames.ContactsFolder,
                typeof(ContactsFolder),
                delegate(ExchangeService srv) { return new ContactsFolder(srv); },
                null);

            // ContactGroup
            this.AddServiceObjectType(
                XmlElementNames.DistributionList,
                typeof(ContactGroup),
                delegate(ExchangeService srv) { return new ContactGroup(srv); },
                delegate(ItemAttachment itemAttachment, bool isNew) { return new ContactGroup(itemAttachment); });

            // Conversation
            this.AddServiceObjectType(
                XmlElementNames.Conversation,
                typeof(Conversation),
                delegate(ExchangeService srv) { return new Conversation(srv); },
                null);

            // EmailMessage
            this.AddServiceObjectType(
                XmlElementNames.Message,
                typeof(EmailMessage),
                delegate(ExchangeService srv) { return new EmailMessage(srv); },
                delegate(ItemAttachment itemAttachment, bool isNew) { return new EmailMessage(itemAttachment); });

            // Folder
            this.AddServiceObjectType(
                XmlElementNames.Folder,
                typeof(Folder),
                delegate(ExchangeService srv) { return new Folder(srv); },
                null);

            // Item
            this.AddServiceObjectType(
                XmlElementNames.Item,
                typeof(Item),
                delegate(ExchangeService srv) { return new Item(srv); },
                delegate(ItemAttachment itemAttachment, bool isNew) { return new Item(itemAttachment); });

            // MeetingCancellation
            this.AddServiceObjectType(
                XmlElementNames.MeetingCancellation,
                typeof(MeetingCancellation),
                delegate(ExchangeService srv) { return new MeetingCancellation(srv); },
                delegate(ItemAttachment itemAttachment, bool isNew) { return new MeetingCancellation(itemAttachment); });

            // MeetingMessage
            this.AddServiceObjectType(
                XmlElementNames.MeetingMessage,
                typeof(MeetingMessage),
                delegate(ExchangeService srv) { return new MeetingMessage(srv); },
                delegate(ItemAttachment itemAttachment, bool isNew) { return new MeetingMessage(itemAttachment); });

            // MeetingRequest
            this.AddServiceObjectType(
                XmlElementNames.MeetingRequest,
                typeof(MeetingRequest),
                delegate(ExchangeService srv) { return new MeetingRequest(srv); },
                delegate(ItemAttachment itemAttachment, bool isNew) { return new MeetingRequest(itemAttachment); });

            // MeetingResponse
            this.AddServiceObjectType(
                XmlElementNames.MeetingResponse,
                typeof(MeetingResponse),
                delegate(ExchangeService srv) { return new MeetingResponse(srv); },
                delegate(ItemAttachment itemAttachment, bool isNew) { return new MeetingResponse(itemAttachment); });

            // Persona
            this.AddServiceObjectType(
                XmlElementNames.Persona,
                typeof(Persona),
                delegate(ExchangeService srv) { return new Persona(srv); },
                null);

            // PostItem
            this.AddServiceObjectType(
                XmlElementNames.PostItem,
                typeof(PostItem),
                delegate(ExchangeService srv) { return new PostItem(srv); },
                delegate(ItemAttachment itemAttachment, bool isNew) { return new PostItem(itemAttachment); });

            // SearchFolder
            this.AddServiceObjectType(
                XmlElementNames.SearchFolder,
                typeof(SearchFolder),
                delegate(ExchangeService srv) { return new SearchFolder(srv); },
                null);

            // Task
            this.AddServiceObjectType(
                XmlElementNames.Task,
                typeof(Task),
                delegate(ExchangeService srv) { return new Task(srv); },
                delegate(ItemAttachment itemAttachment, bool isNew) { return new Task(itemAttachment); });

            // TasksFolder
            this.AddServiceObjectType(
                XmlElementNames.TasksFolder,
                typeof(TasksFolder),
                delegate(ExchangeService srv) { return new TasksFolder(srv); },
                null);
        }

        /// <summary>
        /// Adds specified type of service object to map.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="type">The ServiceObject type.</param>
        /// <param name="createServiceObjectWithServiceParam">Delegate to create service object with service param.</param>
        /// <param name="createServiceObjectWithAttachmentParam">Delegate to create service object with attachment param.</param>
        private void AddServiceObjectType(
            string xmlElementName,
            Type type,
            CreateServiceObjectWithServiceParam createServiceObjectWithServiceParam,
            CreateServiceObjectWithAttachmentParam createServiceObjectWithAttachmentParam)
        {
            this.xmlElementNameToServiceObjectClassMap.Add(xmlElementName, type);
            this.serviceObjectConstructorsWithServiceParam.Add(type, createServiceObjectWithServiceParam);
            if (createServiceObjectWithAttachmentParam != null)
            {
                this.serviceObjectConstructorsWithAttachmentParam.Add(type, createServiceObjectWithAttachmentParam);
            }
        }

        /// <summary>
        /// Return Dictionary that maps from element name to ServiceObject Type.
        /// </summary>
        internal Dictionary<string, Type> XmlElementNameToServiceObjectClassMap
        {
            get { return this.xmlElementNameToServiceObjectClassMap; }
        }

        /// <summary>
        /// Return Dictionary that maps from ServiceObject Type to CreateServiceObjectWithServiceParam delegate with ExchangeService parameter.
        /// </summary>
        internal Dictionary<Type, CreateServiceObjectWithServiceParam> ServiceObjectConstructorsWithServiceParam
        {
            get { return this.serviceObjectConstructorsWithServiceParam; }
        }

        /// <summary>
        /// Return Dictionary that maps from ServiceObject Type to CreateServiceObjectWithAttachmentParam delegate with ItemAttachment parameter.
        /// </summary>
        internal Dictionary<Type, CreateServiceObjectWithAttachmentParam> ServiceObjectConstructorsWithAttachmentParam
        {
            get { return this.serviceObjectConstructorsWithAttachmentParam; }
        }
    }
}