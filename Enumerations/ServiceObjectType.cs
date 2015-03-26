// ---------------------------------------------------------------------------
// <copyright file="ServiceObjectType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ServiceObjectType enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the type of a service object.
    /// </summary>
    public enum ServiceObjectType
    {
        /// <summary>
        /// The object is a folder.
        /// </summary>
        Folder,

        /// <summary>
        /// The object is an item.
        /// </summary>
        Item,

        /// <summary>
        /// Data represents a conversation
        /// </summary>
        Conversation,

        /// <summary>
        /// Data represents a persona
        /// </summary>
        Persona
    }
}
