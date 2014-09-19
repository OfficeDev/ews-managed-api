// ---------------------------------------------------------------------------
// <copyright file="DefaultExtendedPropertySet.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DefaultExtendedPropertySet enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the default sets of extended properties.
    /// </summary>
    public enum DefaultExtendedPropertySet
    {
        /// <summary>
        /// The Meeting extended property set.
        /// </summary>
        Meeting,

        /// <summary>
        /// The Appointment extended property set.
        /// </summary>
        Appointment,

        /// <summary>
        /// The Common extended property set.
        /// </summary>
        Common,

        /// <summary>
        /// The PublicStrings extended property set.
        /// </summary>
        PublicStrings,

        /// <summary>
        /// The Address extended property set.
        /// </summary>
        Address,

        /// <summary>
        /// The InternetHeaders extended property set.
        /// </summary>
        InternetHeaders,

        /// <summary>
        /// The CalendarAssistants extended property set.
        /// </summary>
        CalendarAssistant,

        /// <summary>
        /// The UnifiedMessaging extended property set.
        /// </summary>
        UnifiedMessaging,

        /// <summary>
        /// The Task extended property set.
        /// </summary>
        Task
    }
}
