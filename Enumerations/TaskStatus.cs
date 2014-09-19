// ---------------------------------------------------------------------------
// <copyright file="TaskStatus.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the TaskStatus enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the execution status of a task.
    /// </summary>
    public enum TaskStatus
    {
        /// <summary>
        /// The execution of the task is not started.
        /// </summary>
        NotStarted,

        /// <summary>
        /// The execution of the task is in progress.
        /// </summary>
        InProgress,

        /// <summary>
        /// The execution of the task is completed.
        /// </summary>
        Completed,

        /// <summary>
        /// The execution of the task is waiting on others.
        /// </summary>
        WaitingOnOthers,

        /// <summary>
        /// The execution of the task is deferred.
        /// </summary>
        Deferred
    }
}
