// ---------------------------------------------------------------------------
// <copyright file="InstallAppResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the InstallAppResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.ObjectModel;
    using System.IO;
    using System.Xml;

    /// <summary>
    /// Represents the response to a InstallApp operation.
    /// Today this class doesn't add extra functionality. Keep this class here so future
    /// we can return extension info up-on installation complete. 
    /// </summary>
    internal sealed class InstallAppResponse : ServiceResponse
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="InstallAppResponse"/> class.
        /// </summary>
        public InstallAppResponse()
            : base()
        {
        }
    }
}
