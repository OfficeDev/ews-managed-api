// ---------------------------------------------------------------------------
// <copyright file="SetHoldOnMailboxesParameters.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SetHoldOnMailboxesParameters class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents set hold on mailboxes parameters.
    /// </summary>
    public sealed class SetHoldOnMailboxesParameters
    {
        /// <summary>
        /// Action type
        /// </summary>
        public HoldAction ActionType { get; set; }

        /// <summary>
        /// Hold id
        /// </summary>
        public string HoldId { get; set; }

        /// <summary>
        /// Query
        /// </summary>
        public string Query { get; set; }

        /// <summary>
        /// Collection of mailboxes
        /// </summary>
        public string[] Mailboxes { get; set; }

        /// <summary>
        /// Query language
        /// </summary>
        public string Language { get; set; }

        /// <summary>
        /// In-place hold identity
        /// </summary>
        public string InPlaceHoldIdentity { get; set; }
    }
}