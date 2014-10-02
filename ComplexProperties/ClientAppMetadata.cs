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
// <summary>Defines the ClientAppMetadata class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.IO;
    using System.Xml;

    /// <summary>
    /// Represents a ClientAppMetadata object.
    /// </summary>
    public sealed class ClientAppMetadata : ComplexProperty
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ClientAppMetadata"/> class.
        /// </summary>
        internal ClientAppMetadata()
            : base()
        {
            this.Namespace = XmlNamespace.Types;
        }

        /// <summary>
        /// The End node url for the app.
        /// </summary>
        public string EndNodeUrl
        {
            get;
            private set;
        }

        /// <summary>
        /// The action url for the app.
        /// </summary>
        public string ActionUrl
        {
            get;
            private set;
        }

        /// <summary>
        /// The app status for the app.
        /// </summary>
        public string AppStatus
        {
            get;
            private set;
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
                case XmlElementNames.EndNodeUrl:
                    this.EndNodeUrl = reader.ReadElementValue<string>();
                    return true;
                case XmlElementNames.ActionUrl:
                    this.ActionUrl = reader.ReadElementValue<string>();
                    return true;
                case XmlElementNames.AppStatus:
                    this.AppStatus = reader.ReadElementValue<string>();
                    return true;
                default:
                    return false;
            }
        }
    }
}
