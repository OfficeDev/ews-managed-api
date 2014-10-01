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
// <summary>Defines the ConversationAction class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// ConversationAction class that represents ConversationActionType in the request XML.
    /// This class really is meant for representing single ConversationAction that needs to
    /// be taken on a conversation.
    /// </summary>
    internal class ConversationAction : IJsonSerializable
    {
        /// <summary>
        /// Gets or sets conversation action
        /// </summary>
        internal ConversationActionType Action { get; set; }

        /// <summary>
        /// Gets or sets conversation id
        /// </summary>
        internal ConversationId ConversationId { get; set; }

        /// <summary>
        /// Gets or sets ProcessRightAway
        /// </summary>
        internal bool ProcessRightAway { get; set; }

        /// <summary>
        /// Gets or set conversation categories for Always Categorize action
        /// </summary>
        internal StringList Categories { get; set; }

        /// <summary>
        /// Gets or sets Enable Always Delete value for Always Delete action
        /// </summary>
        internal bool EnableAlwaysDelete { get; set; }

        /// <summary>
        /// Gets or sets the IsRead state.
        /// </summary>
        internal bool? IsRead { get; set; }

        /// <summary>
        /// Gets or sets the SuppressReadReceipts flag.
        /// </summary>
        internal bool? SuppressReadReceipts { get; set; }

        /// <summary>
        /// Gets or sets the Deletion mode.
        /// </summary>
        internal DeleteMode? DeleteType { get; set; }

        /// <summary>
        /// Gets or sets the flag.
        /// </summary>
        internal Flag Flag { get; set; }

        /// <summary>
        /// ConversationLastSyncTime is used in one time action to determine the items
        /// on which to take the action.
        /// </summary>
        internal DateTime? ConversationLastSyncTime { get; set; }

        /// <summary>
        /// Gets or sets folder id ContextFolder
        /// </summary>
        internal FolderIdWrapper ContextFolderId { get; set; }

        /// <summary>
        /// Gets or sets folder id for Move action
        /// </summary>
        internal FolderIdWrapper DestinationFolderId { get; set; }

        /// <summary>
        /// Gets or sets the retention policy type.
        /// </summary>
        internal RetentionType? RetentionPolicyType { get; set; }

        /// <summary>
        /// Gets or sets the retention policy tag id.
        /// </summary>
        internal Guid? RetentionPolicyTagId { get; set; }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal string GetXmlElementName()
        {
            return XmlElementNames.ApplyConversationAction;
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal void Validate()
        {
            EwsUtilities.ValidateParam(this.ConversationId, "conversationId");
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(
                XmlNamespace.Types,
                XmlElementNames.ConversationAction);
            try
            {
                string actionValue = String.Empty;
                switch (this.Action)
                {
                    case ConversationActionType.AlwaysCategorize:
                        actionValue = XmlElementNames.AlwaysCategorize;
                        break;
                    case ConversationActionType.AlwaysDelete:
                        actionValue = XmlElementNames.AlwaysDelete;
                        break;
                    case ConversationActionType.AlwaysMove:
                        actionValue = XmlElementNames.AlwaysMove;
                        break;
                    case ConversationActionType.Delete:
                        actionValue = XmlElementNames.Delete;
                        break;
                    case ConversationActionType.Copy:
                        actionValue = XmlElementNames.Copy;
                        break;
                    case ConversationActionType.Move:
                        actionValue = XmlElementNames.Move;
                        break;
                    case ConversationActionType.SetReadState:
                        actionValue = XmlElementNames.SetReadState;
                        break;
                    case ConversationActionType.SetRetentionPolicy:
                        actionValue = XmlElementNames.SetRetentionPolicy;
                        break;
                    case ConversationActionType.Flag:
                        actionValue = XmlElementNames.Flag;
                        break;
                    default:
                        throw new ArgumentException("ConversationAction");
                }

                // Emit the action element
                writer.WriteElementValue(
                                    XmlNamespace.Types,
                                    XmlElementNames.Action,
                                    actionValue);

                // Emit the conversation id element
                this.ConversationId.WriteToXml(
                                    writer,
                                    XmlNamespace.Types,
                                    XmlElementNames.ConversationId);

                if (this.Action == ConversationActionType.AlwaysCategorize ||
                    this.Action == ConversationActionType.AlwaysDelete ||
                    this.Action == ConversationActionType.AlwaysMove)
                {
                    // Emit the ProcessRightAway element
                    writer.WriteElementValue(
                                        XmlNamespace.Types,
                                        XmlElementNames.ProcessRightAway,
                                        EwsUtilities.BoolToXSBool(this.ProcessRightAway));
                }

                if (this.Action == ConversationActionType.AlwaysCategorize)
                {
                    // Emit the categories element
                    if (this.Categories != null && this.Categories.Count > 0)
                    {
                        this.Categories.WriteToXml(
                                   writer,
                                   XmlNamespace.Types,
                                   XmlElementNames.Categories);
                    }
                }
                else if (this.Action == ConversationActionType.AlwaysDelete)
                {
                    // Emit the EnableAlwaysDelete element
                    writer.WriteElementValue(
                                   XmlNamespace.Types,
                                   XmlElementNames.EnableAlwaysDelete,
                                   EwsUtilities.BoolToXSBool(this.EnableAlwaysDelete));
                }
                else if (this.Action == ConversationActionType.AlwaysMove)
                {
                    // Emit the Move Folder Id
                    if (this.DestinationFolderId != null)
                    {
                        writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.DestinationFolderId);
                        this.DestinationFolderId.WriteToXml(writer);
                        writer.WriteEndElement();
                    }
                }
                else
                {
                    if (this.ContextFolderId != null)
                    {
                        writer.WriteStartElement(
                            XmlNamespace.Types,
                            XmlElementNames.ContextFolderId);

                        this.ContextFolderId.WriteToXml(writer);

                        writer.WriteEndElement();
                    }

                    if (this.ConversationLastSyncTime.HasValue)
                    {
                        writer.WriteElementValue(
                            XmlNamespace.Types,
                            XmlElementNames.ConversationLastSyncTime,
                            this.ConversationLastSyncTime.Value);
                    }

                    if (this.Action == ConversationActionType.Copy)
                    {
                        EwsUtilities.Assert(
                            this.DestinationFolderId != null,
                            "ApplyconversationActionRequest",
                            "DestinationFolderId should be set when performing copy action");

                        writer.WriteStartElement(
                            XmlNamespace.Types,
                            XmlElementNames.DestinationFolderId);
                        this.DestinationFolderId.WriteToXml(writer);
                        writer.WriteEndElement();
                    }
                    else if (this.Action == ConversationActionType.Move)
                    {
                        EwsUtilities.Assert(
                            this.DestinationFolderId != null,
                            "ApplyconversationActionRequest",
                            "DestinationFolderId should be set when performing move action");

                        writer.WriteStartElement(
                            XmlNamespace.Types,
                            XmlElementNames.DestinationFolderId);
                        this.DestinationFolderId.WriteToXml(writer);
                        writer.WriteEndElement();
                    }
                    else if (this.Action == ConversationActionType.Delete)
                    {
                        EwsUtilities.Assert(
                            this.DeleteType.HasValue,
                            "ApplyconversationActionRequest",
                            "DeleteType should be specified when deleting a conversation.");

                        writer.WriteElementValue(
                            XmlNamespace.Types,
                            XmlElementNames.DeleteType,
                            this.DeleteType.Value);
                    }
                    else if (this.Action == ConversationActionType.SetReadState)
                    {
                        EwsUtilities.Assert(
                            this.IsRead.HasValue,
                            "ApplyconversationActionRequest",
                            "IsRead should be specified when marking/unmarking a conversation as read.");

                        writer.WriteElementValue(
                            XmlNamespace.Types,
                            XmlElementNames.IsRead,
                            this.IsRead.Value);

                        if (this.SuppressReadReceipts.HasValue)
                        {
                            writer.WriteElementValue(
                                XmlNamespace.Types,
                                XmlElementNames.SuppressReadReceipts,
                                this.SuppressReadReceipts.Value);
                        }
                    }
                    else if (this.Action == ConversationActionType.SetRetentionPolicy)
                    {
                        EwsUtilities.Assert(
                            this.RetentionPolicyType.HasValue,
                            "ApplyconversationActionRequest",
                            "RetentionPolicyType should be specified when setting a retention policy on a conversation.");

                        writer.WriteElementValue(
                            XmlNamespace.Types,
                            XmlElementNames.RetentionPolicyType,
                            this.RetentionPolicyType.Value);

                        if (this.RetentionPolicyTagId.HasValue)
                        {
                            writer.WriteElementValue(
                                XmlNamespace.Types,
                                XmlElementNames.RetentionPolicyTagId,
                                this.RetentionPolicyTagId.Value);
                        }
                    }
                    else if (this.Action == ConversationActionType.Flag)
                    {
                        EwsUtilities.Assert(
                            this.Flag != null,
                            "ApplyconversationActionRequest",
                            "Flag should be specified when flagging conversation items.");

                        writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Flag);
                        this.Flag.WriteElementsToXml(writer);
                        writer.WriteEndElement();
                    }
                }
            }
            finally
            {
                writer.WriteEndElement();
            }
        }

        /// <summary>
        /// Creates a JSON representation of this object.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        public object ToJson(ExchangeService service)
        {
            JsonObject jsonProperty = new JsonObject();

            // Emit the action element
            jsonProperty.Add(XmlElementNames.Action, this.Action);

            // Emit the conversation id element
            jsonProperty.Add(XmlElementNames.ConversationId, this.ConversationId.InternalToJson(service));

            if (this.Action == ConversationActionType.AlwaysCategorize ||
                this.Action == ConversationActionType.AlwaysDelete ||
                this.Action == ConversationActionType.AlwaysMove)
            {
                // Emit the ProcessRightAway element
                jsonProperty.Add(XmlElementNames.ProcessRightAway, this.ProcessRightAway);
            }

            if (this.Action == ConversationActionType.AlwaysCategorize)
            {
                // Emit the categories element
                if (this.Categories != null && this.Categories.Count > 0)
                {
                    jsonProperty.Add(XmlElementNames.Categories, this.Categories.InternalToJson(service));
                }
            }
            else if (this.Action == ConversationActionType.AlwaysDelete)
            {
                // Emit the EnableAlwaysDelete element
                jsonProperty.Add(XmlElementNames.EnableAlwaysDelete, this.EnableAlwaysDelete);
            }
            else if (this.Action == ConversationActionType.AlwaysMove)
            {
                // Emit the Move Folder Id
                if (this.DestinationFolderId != null)
                {
                    JsonObject jsonTargetFolderId = new JsonObject();
                    jsonTargetFolderId.Add(XmlElementNames.BaseFolderId, this.DestinationFolderId.InternalToJson(service));
                    jsonProperty.Add(XmlElementNames.DestinationFolderId, jsonTargetFolderId);
                }
            }
            else
            {
                if (this.ContextFolderId != null)
                {
                    JsonObject jsonTargetFolderId = new JsonObject();
                    jsonTargetFolderId.Add(XmlElementNames.BaseFolderId, this.ContextFolderId.InternalToJson(service));
                    jsonProperty.Add(XmlElementNames.ContextFolderId, jsonTargetFolderId);
                }

                if (this.ConversationLastSyncTime.HasValue)
                {
                    jsonProperty.Add( XmlElementNames.ConversationLastSyncTime, this.ConversationLastSyncTime.Value);
                }

                if (this.Action == ConversationActionType.Copy)
                {
                    EwsUtilities.Assert(
                        this.DestinationFolderId != null,
                        "ApplyconversationActionRequest",
                        "DestinationFolderId should be set when performing copy action");

                    JsonObject jsonTargetFolderId = new JsonObject();
                    jsonTargetFolderId.Add(XmlElementNames.BaseFolderId, this.DestinationFolderId.InternalToJson(service));
                    jsonProperty.Add(XmlElementNames.DestinationFolderId, jsonTargetFolderId);
                }
                else if (this.Action == ConversationActionType.Move)
                {
                    EwsUtilities.Assert(
                        this.DestinationFolderId != null,
                        "ApplyconversationActionRequest",
                        "DestinationFolderId should be set when performing move action");

                    JsonObject jsonTargetFolderId = new JsonObject();
                    jsonTargetFolderId.Add(XmlElementNames.BaseFolderId, this.DestinationFolderId.InternalToJson(service));
                    jsonProperty.Add(XmlElementNames.DestinationFolderId, jsonTargetFolderId);
                }
                else if (this.Action == ConversationActionType.Delete)
                {
                    EwsUtilities.Assert(
                        this.DeleteType.HasValue,
                        "ApplyconversationActionRequest",
                        "DeleteType should be specified when deleting a conversation.");

                    jsonProperty.Add(XmlElementNames.DeleteType, this.DeleteType.Value);
                }
                else if (this.Action == ConversationActionType.SetReadState)
                {
                    EwsUtilities.Assert(
                        this.IsRead.HasValue,
                        "ApplyconversationActionRequest",
                        "IsRead should be specified when marking/unmarking a conversation as read.");

                    jsonProperty.Add(XmlElementNames.IsRead, this.IsRead.Value);

                    if (this.SuppressReadReceipts.HasValue)
                    {
                        jsonProperty.Add(XmlElementNames.SuppressReadReceipts, this.SuppressReadReceipts.HasValue);
                    }
                }
                else if (this.Action == ConversationActionType.SetRetentionPolicy)
                {
                    EwsUtilities.Assert(
                        this.RetentionPolicyType.HasValue,
                        "ApplyconversationActionRequest",
                        "RetentionPolicyType should be specified when setting a retention policy on a conversation.");

                    jsonProperty.Add(XmlElementNames.RetentionPolicyType, this.RetentionPolicyType.Value);

                    if (this.RetentionPolicyTagId.HasValue)
                    {
                        jsonProperty.Add(XmlElementNames.RetentionPolicyTagId, this.RetentionPolicyTagId.Value);
                    }
                }
                else if (this.Action == ConversationActionType.Flag)
                {
                    EwsUtilities.Assert(
                        this.Flag != null,
                        "ApplyconversationActionRequest",
                        "Flag should be specified when flagging items in a conversation.");

                    jsonProperty.Add(XmlElementNames.Flag, this.Flag.InternalToJson(service));
                }
            } 

            return jsonProperty;
        }
    }
}