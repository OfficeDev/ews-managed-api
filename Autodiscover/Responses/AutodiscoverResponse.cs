// ---------------------------------------------------------------------------
// <copyright file="AutodiscoverResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AutodiscoverResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents the base class for all responses returned by the Autodiscover service.
    /// </summary>
    public abstract class AutodiscoverResponse
    {
        private AutodiscoverErrorCode errorCode;
        private string errorMessage;
        private Uri redirectionUrl;

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverResponse"/> class.
        /// </summary>
        internal AutodiscoverResponse()
        {
            this.errorCode = AutodiscoverErrorCode.NoError;
            this.errorMessage = Strings.NoError;
        }

        /// <summary>
        /// Loads response from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="endElementName">End element name.</param>
        internal virtual void LoadFromXml(EwsXmlReader reader, string endElementName)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.ErrorCode:
                    this.ErrorCode = reader.ReadElementValue<AutodiscoverErrorCode>();
                    break;
                case XmlElementNames.ErrorMessage:
                    this.ErrorMessage = reader.ReadElementValue();
                    break;
                default:
                    break;
            }
        }

        #region Properties
        /// <summary>
        /// Gets the error code that was returned by the service.
        /// </summary>
        public AutodiscoverErrorCode ErrorCode
        { 
            get { return this.errorCode; }
            internal set { this.errorCode = value; }
        }

        /// <summary>
        /// Gets the error message that was returned by the service.
        /// </summary>
        /// <value>The error message.</value>
        public string ErrorMessage
        {
            get { return this.errorMessage; }
            internal set { this.errorMessage = value; }
        }

        /// <summary>
        /// Gets or sets the redirection URL.
        /// </summary>
        /// <value>The redirection URL.</value>
        internal Uri RedirectionUrl
        {
            get { return this.redirectionUrl; }
            set { this.redirectionUrl = value; }
        }

        #endregion
    }
}
