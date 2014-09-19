// ---------------------------------------------------------------------------
// <copyright file="ConversationQueryTraversal.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ConversationQueryTraversal enumeration.</summary>
//-----------------------------------------------------------------------

using System.Xml.Serialization;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the folder traversal depth in queries.
    /// </summary>
    public enum ConversationQueryTraversal
    {
        /// <summary>
        /// Shallow traversal
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        Shallow = 0,

        /// <summary>
        /// Deep traversal
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        Deep = 1,
    }
}