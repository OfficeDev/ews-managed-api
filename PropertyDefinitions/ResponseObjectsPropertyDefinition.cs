// ---------------------------------------------------------------------------
// <copyright file="ResponseObjectsPropertyDefinition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ResponseObjectsPropertyDefinition class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Represents response object property defintion.
    /// </summary>
    internal sealed class ResponseObjectsPropertyDefinition : PropertyDefinition
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ResponseObjectsPropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="version">The version.</param>
        internal ResponseObjectsPropertyDefinition(
            string xmlElementName,
            string uri,
            ExchangeVersion version)
            : base(
                xmlElementName,
                uri,
                version)
        {
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="propertyBag">The property bag.</param>
        internal override sealed void LoadPropertyValueFromXml(EwsServiceXmlReader reader, PropertyBag propertyBag)
        {
            ResponseActions value = ResponseActions.None;

            reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, this.XmlElementName);

            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if (reader.IsStartElement())
                    {
                        value |= GetResponseAction(reader.LocalName);
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Types, this.XmlElementName));
            }

            propertyBag[this] = value;
        }

        /// <summary>
        /// Loads the property value from json.
        /// </summary>
        /// <param name="value">The JSON value.  Can be a JsonObject, string, number, bool, array, or null.</param>
        /// <param name="service">The service.</param>
        /// <param name="propertyBag">The property bag.</param>
        /// <remarks>
        /// The ResponseActions collection is returned as an array of values of derived ResponseObject types. For example:
        /// "ResponseObjects" : [ { "__type" : "CancelCalendarItem:#Exchange" }, { "__type" : "ForwardItem:#Exchange" } ]
        /// </remarks>
        internal override void LoadPropertyValueFromJson(object value, ExchangeService service, PropertyBag propertyBag)
        {
            ResponseActions responseActionValue = ResponseActions.None;

            object[] jsonResponseActions = value as object[];

            if (jsonResponseActions != null)
            {
                foreach (JsonObject jsonResponseAction in jsonResponseActions.OfType<JsonObject>())
                {
                    if (jsonResponseAction.HasTypeProperty())
                    {
                        string actionString = jsonResponseAction.ReadTypeString();

                        if (!string.IsNullOrEmpty(actionString))
                        {
                            responseActionValue |= GetResponseAction(actionString);
                        }
                    }
                }
            }

            propertyBag[this] = responseActionValue;
        }

        /// <summary>
        /// Gets the response action.
        /// </summary>
        /// <param name="responseActionString">The response action string.</param>
        /// <returns></returns>
        private static ResponseActions GetResponseAction(string responseActionString)
        {
            ResponseActions value = ResponseActions.None;

            switch (responseActionString)
            {
                case XmlElementNames.AcceptItem:
                    value = ResponseActions.Accept;
                    break;
                case XmlElementNames.TentativelyAcceptItem:
                    value = ResponseActions.TentativelyAccept;
                    break;
                case XmlElementNames.DeclineItem:
                    value = ResponseActions.Decline;
                    break;
                case XmlElementNames.ReplyToItem:
                    value = ResponseActions.Reply;
                    break;
                case XmlElementNames.ForwardItem:
                    value = ResponseActions.Forward;
                    break;
                case XmlElementNames.ReplyAllToItem:
                    value = ResponseActions.ReplyAll;
                    break;
                case XmlElementNames.CancelCalendarItem:
                    value = ResponseActions.Cancel;
                    break;
                case XmlElementNames.RemoveItem:
                    value = ResponseActions.RemoveFromCalendar;
                    break;
                case XmlElementNames.SuppressReadReceipt:
                    value = ResponseActions.SuppressReadReceipt;
                    break;
                case XmlElementNames.PostReplyItem:
                    value = ResponseActions.PostReply;
                    break;
            }
            return value;
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="propertyBag">The property bag.</param>
        /// <param name="isUpdateOperation">Indicates whether the context is an update operation.</param>
        internal override void WritePropertyValueToXml(
            EwsServiceXmlWriter writer,
            PropertyBag propertyBag,
            bool isUpdateOperation)
        {
            // ResponseObjects is a read-only property, no need to implement this.
        }

        /// <summary>
        /// Writes the json value.
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        /// <param name="propertyBag">The property bag.</param>
        /// <param name="service">The service.</param>
        /// <param name="isUpdateOperation">if set to <c>true</c> [is update operation].</param>
        internal override void WriteJsonValue(JsonObject jsonObject, PropertyBag propertyBag, ExchangeService service, bool isUpdateOperation)
        {
            // ResponseObjects is a read-only property, no need to implement this.
        }

        /// <summary>
        /// Gets a value indicating whether this property definition is for a nullable type (ref, int?, bool?...).
        /// </summary>
        internal override bool IsNullable
        {
            get { return false; }
        }

        /// <summary>
        /// Gets the property type.
        /// </summary>
        public override Type Type
        {
            get { return typeof(ResponseActions); }
        }
    }
}