// ---------------------------------------------------------------------------
// <copyright file="AccountIsLockedException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AccountIsLockedException class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents an error that occurs when the account that is being accessed is locked and requires user interaction to be unlocked.
    /// </summary>
    [Serializable]
    public class AccountIsLockedException : ServiceRemoteException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AccountIsLockedException"/> class.
        /// </summary>
        /// <param name="message">Error message text.</param>
        /// <param name="accountUnlockUrl">URL for client to visit to unlock account.</param>
        /// <param name="innerException">Inner exception.</param>
        public AccountIsLockedException(string message, Uri accountUnlockUrl, Exception innerException)
            : base(message, innerException)
        {
            this.AccountUnlockUrl = accountUnlockUrl;
        }

        /// <summary>
        /// Gets the URL of a web page where the user can navigate to unlock his or her account.
        /// </summary>
        public Uri AccountUnlockUrl
        {
            get;
            private set;
        }
    }
}
