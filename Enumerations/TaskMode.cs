// ---------------------------------------------------------------------------
// <copyright file="TaskMode.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the TaskMode enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the modes of a Task.
    /// </summary>
    public enum TaskMode
    {
        /// <summary>
        /// The task is normal
        /// </summary>
        Normal = 0,

        /// <summary>
        /// The task is a task assignment request
        /// </summary>
        Request = 1,

        /// <summary>
        /// The task assignment request was accepted
        /// </summary>
        RequestAccepted = 2,

        /// <summary>
        /// The task assignment request was declined
        /// </summary>
        RequestDeclined = 3,

        /// <summary>
        /// The task has been updated
        /// </summary>
        Update = 4,

        /// <summary>
        /// The task is self delegated
        /// </summary>
        SelfDelegated = 5
    }
}
