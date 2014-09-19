// ---------------------------------------------------------------------------
// <copyright file="DelegateManagementRequestBase.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DelegateManagementRequestBase class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an abstract delegate management request.
    /// </summary>
    /// <typeparam name="TResponse">The type of the response.</typeparam>
    internal abstract class DelegateManagementRequestBase<TResponse> : SimpleServiceRequestBase
        where TResponse : DelegateManagementResponse
    {
        private Mailbox mailbox;

        /// <summary>
        /// Initializes a new instance of the <see cref="DelegateManagementRequestBase&lt;TResponse&gt;"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal DelegateManagementRequestBase(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParam(this.Mailbox, "Mailbox");
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            this.Mailbox.WriteToXml(
                writer,
                XmlNamespace.Messages,
                XmlElementNames.Mailbox);
        }

        /// <summary>
        /// Creates the response.
        /// </summary>
        /// <returns>Response object.</returns>
        internal abstract TResponse CreateResponse();

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            DelegateManagementResponse response = this.CreateResponse();

            response.LoadFromXml(reader, this.GetResponseXmlElementName());

            return response;
        }

        /// <summary>
        /// Executes this request.
        /// </summary>
        /// <returns>Response object.</returns>
        internal TResponse Execute()
        {
            TResponse serviceResponse = (TResponse)this.InternalExecute();

            serviceResponse.ThrowIfNecessary();

            return serviceResponse;
        }

        /// <summary>
        /// Gets or sets the mailbox.
        /// </summary>
        /// <value>The mailbox.</value>
        public Mailbox Mailbox
        {
            get { return this.mailbox; }
            set { this.mailbox = value; }
        }
    }
}
