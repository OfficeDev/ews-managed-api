// ---------------------------------------------------------------------------
// <copyright file="ProtocolConnectionCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ProtocolConnectionCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System.Collections.Generic;
    using System.Xml;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents a user setting that is a collection of protocol connection.
    /// </summary>
    public sealed class ProtocolConnectionCollection
    {
        private List<ProtocolConnection> connections;

        /// <summary>
        /// Initializes a new instance of the <see cref="ProtocolConnectionCollection"/> class.
        /// </summary>
        internal ProtocolConnectionCollection()
        {
            this.connections = new List<ProtocolConnection>();
        }

        /// <summary>
        /// Read user setting with ProtocolConnectionCollection value.
        /// </summary>
        /// <param name="reader">EwsServiceXmlReader</param>
        internal static ProtocolConnectionCollection LoadFromXml(EwsXmlReader reader)
        {
            ProtocolConnectionCollection value = new ProtocolConnectionCollection();
            ProtocolConnection connection = null;

            do
            {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (reader.LocalName == XmlElementNames.ProtocolConnection)
                    {
                        connection = ProtocolConnection.LoadFromXml(reader);
                        if (connection != null)
                        {
                            value.Connections.Add(connection);
                        }
                    }
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.ProtocolConnections));

            return value;
        }

        /// <summary>
        /// Gets the Connections.
        /// </summary>
        public List<ProtocolConnection> Connections
        {
            get { return this.connections; }
            internal set { this.connections = value; }
        }
    }
}
