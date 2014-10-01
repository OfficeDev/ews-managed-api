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
// <summary>Defines the AlternateMailbox class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System.Xml;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents an alternate mailbox.
    /// </summary>
    public sealed class AlternateMailbox
    {
        private string type;
        private string displayName;
        private string legacyDN;
        private string server;
        private string smtpAddress;
        private string ownerSmtpAddress;

        /// <summary>
        /// Initializes a new instance of the <see cref="AlternateMailbox"/> class.
        /// </summary>
        private AlternateMailbox()
        {
        }

        /// <summary>
        /// Loads AlternateMailbox instance from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>AlternateMailbox.</returns>
        internal static AlternateMailbox LoadFromXml(EwsXmlReader reader)
        {
            AlternateMailbox altMailbox = new AlternateMailbox();

            do
            {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case XmlElementNames.Type:
                            altMailbox.Type = reader.ReadElementValue<string>();
                            break;
                        case XmlElementNames.DisplayName:
                            altMailbox.DisplayName = reader.ReadElementValue<string>();
                            break;
                        case XmlElementNames.LegacyDN:
                            altMailbox.LegacyDN = reader.ReadElementValue<string>();
                            break;
                        case XmlElementNames.Server:
                            altMailbox.Server = reader.ReadElementValue<string>();
                            break;
                        case XmlElementNames.SmtpAddress:
                            altMailbox.SmtpAddress = reader.ReadElementValue<string>();
                            break;
                        case XmlElementNames.OwnerSmtpAddress:
                            altMailbox.OwnerSmtpAddress = reader.ReadElementValue<string>();
                            break;
                        default:
                            break;
                    }
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.AlternateMailbox));

            return altMailbox;
        }

        /// <summary>
        /// Gets the alternate mailbox type.
        /// </summary>
        /// <value>The type.</value>
        public string Type
        {
            get { return this.type; }
            internal set { this.type = value; }
        }

        /// <summary>
        /// Gets the alternate mailbox display name.
        /// </summary>
        public string DisplayName
        {
            get { return this.displayName; }
            internal set { this.displayName = value; }
        }

        /// <summary>
        /// Gets the alternate mailbox legacy DN.
        /// </summary>
        public string LegacyDN
        {
            get { return this.legacyDN; }
            internal set { this.legacyDN = value; }
        }

        /// <summary>
        /// Gets the alernate mailbox server.
        /// </summary>
        public string Server
        {
            get { return this.server; }
            internal set { this.server = value; }
        }

        /// <summary>
        /// Gets the alternate mailbox address.
        /// It has value only when Server and LegacyDN is empty.
        /// </summary>
        public string SmtpAddress
        {
            get { return this.smtpAddress; }
            internal set { this.smtpAddress = value; }
        }

        /// <summary>
        /// Gets the alternate mailbox owner SmtpAddress.
        /// </summary>
        public string OwnerSmtpAddress
        {
            get { return this.ownerSmtpAddress; }
            internal set { this.ownerSmtpAddress = value; }
        }
    }
}
