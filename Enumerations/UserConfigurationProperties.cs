// ---------------------------------------------------------------------------
// <copyright file="UserConfigurationProperties.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the UserConfigurationProperties enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Identifies the user configuration properties to retrieve.
    /// </summary>
    [Flags]
    public enum UserConfigurationProperties
    {
        /// <summary>
        /// Retrieve the Id property.
        /// </summary>
        Id = 1,

        /// <summary>
        /// Retrieve the Dictionary property.
        /// </summary>
        Dictionary = 2,

        /// <summary>
        /// Retrieve the XmlData property.
        /// </summary>
        XmlData = 4,

        /// <summary>
        /// Retrieve the BinaryData property.
        /// </summary>
        BinaryData = 8,

        /// <summary>
        /// Retrieve all properties.
        /// </summary>
        All = UserConfigurationProperties.Id | 
            UserConfigurationProperties.Dictionary | 
            UserConfigurationProperties.XmlData | 
            UserConfigurationProperties.BinaryData
    }
}
