// ---------------------------------------------------------------------------
// <copyright file="AppointmentType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AppointmentType enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the type of an appointment.
    /// </summary>
    public enum AppointmentType
    {
        /// <summary>
        /// The appointment is non-recurring.
        /// </summary>
        Single,

        /// <summary>
        /// The appointment is an occurrence of a recurring appointment.
        /// </summary>
        Occurrence,

        /// <summary>
        /// The appointment is an exception of a recurring appointment.
        /// </summary>
        Exception,

        /// <summary>
        /// The appointment is the recurring master of a series.
        /// </summary>
        RecurringMaster
    }
}
