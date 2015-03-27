// ---------------------------------------------------------------------------
// <copyright file="RegisterConsentResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.ObjectModel;
    using System.IO;
    using System.Xml;

    /// <summary>
    /// Represents the response to a RegisterResponse operation.
    /// Today this class doesn't add extra functionality. Keep this class here so in the future
    /// we can return extension info upon installation complete. 
    /// </summary>
    internal sealed class RegisterConsentResponse : ServiceResponse
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RegisterConsentResponse"/> class.
        /// </summary>
        public RegisterConsentResponse()
            : base()
        {
        }
    }
}
