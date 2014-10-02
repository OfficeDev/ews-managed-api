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
// <summary>Defines the UpdateFolderRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an UpdateFolder request.
    /// </summary>
    internal sealed class UpdateFolderRequest : MultiResponseServiceRequest<ServiceResponse>, IJsonSerializable
    {
        private List<Folder> folders = new List<Folder>();

        /// <summary>
        /// Initializes a new instance of the <see cref="UpdateFolderRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal UpdateFolderRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Validates the request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParamCollection(this.Folders, "Folders");
            for (int i = 0; i < this.Folders.Count; i++)
            {
                Folder folder = this.Folders[i];

                if ((folder == null) || folder.IsNew)
                {
                    throw new ArgumentException(string.Format(Strings.FolderToUpdateCannotBeNullOrNew, i));
                }

                folder.Validate();
            }
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="session">The session.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal override ServiceResponse CreateServiceResponse(ExchangeService session, int responseIndex)
        {
            return new UpdateFolderResponse(this.Folders[responseIndex]);
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>Xml element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.UpdateFolder;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>Xml element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.UpdateFolderResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>Xml element name.</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.UpdateFolderResponseMessage;
        }

        /// <summary>
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of expected response messages.</returns>
        internal override int GetExpectedResponseMessageCount()
        {
            return this.folders.Count;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.FolderChanges);

            foreach (Folder folder in this.folders)
            {
                folder.WriteToXmlForUpdate(writer);
            }

            writer.WriteEndElement();
        }

        /// <summary>
        /// Creates a JSON representation of this object.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        object IJsonSerializable.ToJson(ExchangeService service)
        {
            JsonObject jsonRequest = new JsonObject();
            List<object> jsonUpdates = new List<object>();

            foreach (Folder folder in this.folders)
            {
                jsonUpdates.Add(folder.WriteToJsonForUpdate(service));
            }

            jsonRequest.Add(XmlElementNames.FolderChanges, jsonUpdates.ToArray());

            return jsonRequest;
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2007_SP1;
        }

        /// <summary>
        /// Gets the list of folders.
        /// </summary>
        /// <value>The folders.</value>
        public List<Folder> Folders
        {
            get { return this.folders; }
        }
    }
}
