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
// <summary>Defines the DelegateUserResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents the response to an individual delegate user manipulation (add, remove, update) operation.
    /// </summary>
    public sealed class DelegateUserResponse : ServiceResponse
    {
        private bool readDelegateUser;
        private DelegateUser delegateUser;

        /// <summary>
        /// Initializes a new instance of the <see cref="DelegateUserResponse"/> class.
        /// </summary>
        /// <param name="readDelegateUser">if set to <c>true</c> [read delegate user].</param>
        /// <param name="delegateUser">Existing DelegateUser to use (may be null).</param>
        internal DelegateUserResponse(bool readDelegateUser, DelegateUser delegateUser)
            : base()
        {
            this.readDelegateUser = readDelegateUser;
            this.delegateUser = delegateUser;
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            if (this.readDelegateUser)
            {
                if (this.delegateUser == null)
                {
                    this.delegateUser = new DelegateUser();
                }

                reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.DelegateUser);

                this.delegateUser.LoadFromXml(
                    reader,
                    XmlNamespace.Messages,
                    reader.LocalName);
            }
        }

        /// <summary>
        /// The delegate user that was involved in the operation.
        /// </summary>
        public DelegateUser DelegateUser
        {
            get { return this.delegateUser; }
        }
    }
}
