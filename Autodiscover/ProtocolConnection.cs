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
// <summary>Defines the ProtocolConnection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System.Xml;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents the email Protocol connection settings for pop/imap/smtp protocols.
    /// </summary>
    public sealed class ProtocolConnection
    {
        private string encryptionMethod;
        private string hostname;
        private int port;

        /// <summary>
        /// Initializes a new instance of the <see cref="ProtocolConnection"/> class.
        /// </summary>
        internal ProtocolConnection()
        {
        }

        /// <summary>
        /// Read user setting with ProtocolConnection value.
        /// </summary>
        /// <param name="reader">EwsServiceXmlReader</param>
        internal static ProtocolConnection LoadFromXml(EwsXmlReader reader)
        {
            ProtocolConnection connection = new ProtocolConnection();
            
            do
            {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case XmlElementNames.EncryptionMethod:
                            connection.EncryptionMethod = reader.ReadElementValue<string>();
                            break;
                        case XmlElementNames.Hostname:
                            connection.Hostname = reader.ReadElementValue<string>();
                            break;
                        case XmlElementNames.Port:
                            connection.Port = reader.ReadElementValue<int>();
                            break;
                    }
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.ProtocolConnection));

            return connection;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ProtocolConnection"/> class.
        /// </summary>
        /// <param name="encryptionMethod">The encryption method.</param>
        /// <param name="hostname">The hostname.</param>
        /// <param name="port">The port number to use for the portocol.</param>
        internal ProtocolConnection(string encryptionMethod, string hostname, int port)
        {
            this.encryptionMethod = encryptionMethod;
            this.hostname = hostname;
            this.port = port;
        }

        /// <summary>
        /// Gets or sets the encryption method.
        /// </summary>
        /// <value>The encryption method.</value>
        public string EncryptionMethod
        {
            get { return this.encryptionMethod; }
            set { this.encryptionMethod = value; }
        }

        /// <summary>
        /// Gets or sets the Hostname.
        /// </summary>
        /// <value>The hostname.</value>
        public string Hostname
        {
            get { return this.hostname; }
            set { this.hostname = value; }
        }

        /// <summary>
        /// Gets or sets the port number.
        /// </summary>
        /// <value>The port number.</value>
        public int Port
        {
            get { return this.port; }
            set { this.port = value; }
        }
    }
}
