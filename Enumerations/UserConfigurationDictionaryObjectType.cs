// ---------------------------------------------------------------------------
// <copyright file="UserConfigurationDictionaryObjectType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the UserConfigurationDictionaryObjectType enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Identifies the user configuration dictionary key and value types.
    /// </summary>
    public enum UserConfigurationDictionaryObjectType
    {
        /// <summary>
        /// DateTime type.
        /// </summary>
        DateTime,

        /// <summary>
        /// Boolean type.
        /// </summary>
        Boolean,

        /// <summary>
        /// Byte type.
        /// </summary>
        Byte,

        /// <summary>
        /// String type.
        /// </summary>
        String,

        /// <summary>
        /// 32-bit integer type.
        /// </summary>
        Integer32,

        /// <summary>
        /// 32-bit unsigned integer type.
        /// </summary>
        UnsignedInteger32,

        /// <summary>
        /// 64-bit integer type.
        /// </summary>
        Integer64,

        /// <summary>
        /// 64-bit unsigned integer type.
        /// </summary>
        UnsignedInteger64,

        /// <summary>
        /// String array type.
        /// </summary>
        StringArray,

        /// <summary>
        /// Byte array type
        /// </summary>
        ByteArray,
    }
}
