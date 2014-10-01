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
