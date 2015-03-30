/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

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