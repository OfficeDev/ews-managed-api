// ---------------------------------------------------------------------------
// <copyright file="WSSecurityUtilityIdSignedXml.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Security.Cryptography.Xml;
    using System.Threading;
    using System.Xml;

    /// <summary>
    /// A wrapper class to facilitate creating XML signatures around wsu:Id.
    /// </summary>
    internal class WSSecurityUtilityIdSignedXml : SignedXml
    {
        private static long nextId = 0;
        private static string commonPrefix = "uuid-" + Guid.NewGuid().ToString() + "-";

        private XmlDocument document;
        private Dictionary<string, XmlElement> ids;

        /// <summary>
        /// Initializes a new instance of the WSSecurityUtilityIdSignedXml class from the specified XML document. 
        /// </summary>
        /// <param name="document">Xml document.</param>
        public WSSecurityUtilityIdSignedXml(XmlDocument document)
            : base(document)
        {
            this.document = document;
            this.ids = new Dictionary<string, XmlElement>();
        }

        /// <summary>
        /// Get unique Id.
        /// </summary>
        /// <returns>The wsu id.</returns>
        public static string GetUniqueId()
        {
            return commonPrefix + Interlocked.Increment(ref nextId).ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Add the node as reference. 
        /// </summary>
        /// <param name="xpath">The XPath string.</param>
        public void AddReference(string xpath)
        {
            XmlElement element = this.document.SelectSingleNode(
                xpath, 
                WSSecurityBasedCredentials.NamespaceManager) as XmlElement;

            // for now, ignore the error if the node is not found. 
            // EWS may want to sign extra header while such header is never present in autodiscover request.
            // but currently Credentials are unaware of the service type.
            // 
            if (element != null)
            {
                string wsuId = GetUniqueId();

                XmlAttribute wsuIdAttribute = document.CreateAttribute(
                    EwsUtilities.WSSecurityUtilityNamespacePrefix, 
                    "Id", 
                    EwsUtilities.WSSecurityUtilityNamespace);

                wsuIdAttribute.Value = wsuId;
                element.Attributes.Append(wsuIdAttribute);

                Reference reference = new Reference();
                reference.Uri = "#" + wsuId;
                reference.AddTransform(new XmlDsigExcC14NTransform());

                this.AddReference(reference);
                this.ids.Add(wsuId, element);
            }
        }

        /// <summary>
        /// Returns the XmlElement  object with the specified ID from the specified XmlDocument  object.
        /// </summary>
        /// <param name="document">The XmlDocument object to retrieve the XmlElement object from</param>
        /// <param name="idValue">The ID of the XmlElement object to retrieve from the XmlDocument object.</param>
        /// <returns>The XmlElement object with the specified ID from the specified XmlDocument object</returns>
        public override XmlElement GetIdElement(XmlDocument document, string idValue)
        {
            return this.ids[idValue];
        }
    }
}
