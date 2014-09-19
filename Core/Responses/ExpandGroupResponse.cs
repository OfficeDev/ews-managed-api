// ---------------------------------------------------------------------------
// <copyright file="ExpandGroupResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ExpandGroupResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents the response to a group expansion operation.
    /// </summary>
    internal sealed class ExpandGroupResponse : ServiceResponse
    {
        /// <summary>
        /// AD or store group members.
        /// </summary>
        private ExpandGroupResults members = new ExpandGroupResults();

        /// <summary>
        /// Initializes a new instance of the <see cref="ExpandGroupResponse"/> class.
        /// </summary>
        internal ExpandGroupResponse()
            : base()
        {
        }

        /// <summary>
        /// Gets a list of the group's members.
        /// </summary>
        public ExpandGroupResults Members
        {
            get
            {
                return this.members;
            }
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            this.Members.LoadFromXml(reader);
        }
    }
}
