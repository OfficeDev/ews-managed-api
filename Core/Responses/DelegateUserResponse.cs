// ---------------------------------------------------------------------------
// <copyright file="DelegateUserResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

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
