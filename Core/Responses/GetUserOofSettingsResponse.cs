// ---------------------------------------------------------------------------
// <copyright file="GetUserOofSettingsResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetUserOofSettingsResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents response to GetUserOofSettings request.
    /// </summary>
    internal sealed class GetUserOofSettingsResponse : ServiceResponse
    {
        private OofSettings oofSettings;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetUserOofSettingsResponse"/> class.
        /// </summary>
        internal GetUserOofSettingsResponse()
            : base()
        {
        }

        /// <summary>
        /// Gets or sets the OOF settings.
        /// </summary>
        /// <value>The oof settings.</value>
        public OofSettings OofSettings
        {
            get { return this.oofSettings; }
            internal set { this.oofSettings = value; }
        }
    }
}
