// ---------------------------------------------------------------------------
// <copyright file="GetUserRetentionPolicyTagsResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetUserRetentionPolicyTagsResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the GetUserRetentionPolicyTagsResponse response.
    /// </summary>
    public sealed class GetUserRetentionPolicyTagsResponse : ServiceResponse
    {
        List<RetentionPolicyTag> retentionPolicyTags = new List<RetentionPolicyTag>();

        /// <summary>
        /// Initializes a new instance of the <see cref="GetUserRetentionPolicyTagsResponse"/> class.
        /// </summary>
        internal GetUserRetentionPolicyTagsResponse()
            : base()
        {
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            this.retentionPolicyTags.Clear();

            base.ReadElementsFromXml(reader);

            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.RetentionPolicyTags);
            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();
                    if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.RetentionPolicyTag))
                    {
                        this.retentionPolicyTags.Add(RetentionPolicyTag.LoadFromXml(reader));
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.RetentionPolicyTags));
                reader.ReadEndElementIfNecessary(XmlNamespace.Messages, XmlElementNames.RetentionPolicyTags);
            }
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            this.retentionPolicyTags.Clear();

            base.ReadElementsFromJson(responseObject, service);

            if (responseObject.ContainsKey(XmlElementNames.RetentionPolicyTags))
            {
                foreach (object retentionPolicyTagObject in responseObject.ReadAsArray(XmlElementNames.RetentionPolicyTags))
                {
                    JsonObject jsonRetentionPolicyTag = retentionPolicyTagObject as JsonObject;
                    this.retentionPolicyTags.Add(RetentionPolicyTag.LoadFromJson(jsonRetentionPolicyTag));
                }
            }
        }

        /// <summary>
        /// Retention policy tags result.
        /// </summary>
        public RetentionPolicyTag[] RetentionPolicyTags
        {
            get { return this.retentionPolicyTags.ToArray(); }
        }
    }
}
